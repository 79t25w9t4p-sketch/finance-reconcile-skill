#!/usr/bin/env node
'use strict';
const fs=require('node:fs'),path=require('node:path'),cp=require('node:child_process');
const R=require('./core.cjs');
const crypto=require('node:crypto');
const VERSION='0.5.0';
function fingerprint(file){return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');}
function main(){
 const opts={};const argv=process.argv.slice(2);
 if(argv.length===0||argv.includes('--help')){console.log('财务明细核对 '+VERSION+'\n用法: node scripts/reconcile.cjs --left 文件 --right 文件 --inspect\n核对时另加: --left-key 列名 --left-amount 列名 --right-key 列名 --right-amount 列名 --out 新目录\n联合匹配: --left-keys JSON数组 --right-keys JSON数组（两侧字段顺序对应）\nXLSX可选: --left-sheet 名称 --right-sheet 名称 --left-header-row 行号 --right-header-row 行号 --python Python路径\n输出: CSV差异表、完整表、JSON记录、摘要和运行清单。CSV须UTF-8。');return;}
 for(let i=0;i<argv.length;i++){
  const k=argv[i];if(!/^--[a-z-]+$/.test(k)||Object.hasOwn(opts,k))throw Error('Invalid or duplicate option: '+k);
  if(k==='--inspect'||k==='--single-key-confirmed'){opts[k]=true;continue;}
  if(!argv[i+1]||argv[i+1].startsWith('--'))throw Error('Missing value: '+k);
  opts[k]=argv[++i];
 }
 const allowed=['--single-key-confirmed','--left-keys','--right-keys','--inspect','--left','--right','--left-key','--right-key','--left-amount','--right-amount','--left-sheet','--right-sheet','--left-header-row','--right-header-row','--python','--out'];
 for(const k of Object.keys(opts))if(!allowed.includes(k))throw Error('Unknown option: '+k);
 const inputs={};
 function load(side){
  const file=opts['--'+side];if(!file)throw Error('Required: --'+side);
  const ext=path.extname(file).toLowerCase();let content,metadata;
  const header=opts['--'+side+'-header-row']||'1';
  if(!/^[1-9]\d*$/.test(header))throw Error('Header row must be a positive integer');
  if(ext!=='.xlsx'&&header!=='1')throw Error('Custom header row currently supports XLSX only');
  if(fs.statSync(file).size>10*1024*1024)throw Error('Input exceeds 10 MB');
  const before=fingerprint(file);
  if(ext==='.xlsx'){
   const r=cp.spawnSync(opts['--python']||'python3',[path.join(__dirname,'xlsx_read.py'),file,opts['--'+side+'-sheet']||'',header],{encoding:'utf8',maxBuffer:3*1024*1024,timeout:30000});
   if(r.error||r.status!==0)throw Error(r.error?.message||r.stderr.trim());metadata=JSON.parse(r.stdout);content=metadata.csv;
  }else if(['.csv','.tsv'].includes(ext)){
   if(fs.statSync(file).size>2*1024*1024)throw Error('CSV exceeds 2 MB');
   content=new TextDecoder('utf-8',{fatal:true}).decode(fs.readFileSync(file));
  }else throw Error('Supported: UTF-8 CSV, TSV, XLSX; convert XLS first');
  const data=R.parse(content);
  if(metadata){if(metadata.sourceLines.length!==data.rows.length)throw Error('XLSX source row count mismatch');data.sourceLines=metadata.sourceLines;}
  const after=fingerprint(file);if(before!==after)throw Error('Input changed during reading: '+side);
  inputs[side]={file:path.basename(file),sha256:before,sheet:metadata?.sheet||null,headerRow:Number(header),rows:data.rows.length};
  return data;
 }
 const left=load('left'),right=load('right');
 if(opts['--inspect']){console.log(JSON.stringify({left:{headers:left.headers,rows:left.rows.length},right:{headers:right.headers,rows:right.rows.length}},null,2));return;}
 function col(side,type,data){const name=opts['--'+side+'-'+type];const i=data.headers.indexOf(name);if(i<0)throw Error('Missing/unknown --'+side+'-'+type);return i;}
 function keys(side,data){
  const multiple=opts['--'+side+'-keys'];
  if(multiple&&opts['--'+side+'-key'])throw Error('Use either --'+side+'-key or --'+side+'-keys');
  const names=multiple?JSON.parse(multiple):[opts['--'+side+'-key']];
  if(!Array.isArray(names)||!names.length||names.some(x=>typeof x!=='string'||!data.headers.includes(x)))throw Error('Keys must be a JSON array of existing column names');
  return names.map(x=>data.headers.indexOf(x));
 }
 const lk=keys('left',left),rk=keys('right',right);
 if(lk.length!==rk.length)throw Error('Left and right keys must have the same number of fields in corresponding order');
 const scopeNames=['客户','客户名称','公司','公司名称','客户编码','公司编码'];
 if(lk.length===1&&!opts['--single-key-confirmed']&&left.headers.some(x=>scopeNames.includes(x))&&right.headers.some(x=>scopeNames.includes(x)))throw Error('两表包含客户或公司字段，请用 --left-keys / --right-keys 联合匹配；只有已确认单号跨客户全局唯一时才加 --single-key-confirmed');
 const report=R.reconcile(left,right,lk,col('left','amount',left),rk,col('right','amount',right));
 const coverage={};
 for(const side of ['left','right']){
  const records=report.result.flatMap(x=>x[side]);
  const expected=side==='left'?left.rows.length:right.rows.length;
  if(records.length!==expected||new Set(records.map(x=>x.line)).size!==expected)throw Error('Record coverage check failed: '+side);
  coverage[side]={inputRows:expected,reportedRows:records.length,everyRecordExactlyOnce:true};
  if(fingerprint(opts['--'+side])!==inputs[side].sha256)throw Error('Input changed before reporting: '+side);
 }
 const out=opts['--out'];if(!out)throw Error('Required: --out (a new directory)');
 // New directory only: no accidental overwriting of source files or prior reports.
 fs.mkdirSync(out,{recursive:false});
 const counts={};for(const x of report.result)counts[x.status]=(counts[x.status]||0)+1;
 fs.writeFileSync(path.join(out,'差异表.csv'),R.exportCSV(report.result.filter(x=>x.status!=='一致')),{flag:'wx'});
 fs.writeFileSync(path.join(out,'完整核对表.csv'),R.exportCSV(report.result),{flag:'wx'});
 fs.writeFileSync(path.join(out,'核对记录.json'),JSON.stringify(report,(_,v)=>typeof v==='bigint'?v.toString():v,2),{flag:'wx'});
 const summary={leftRows:report.leftRows,rightRows:report.rightRows,leftTotal:R.fmt(report.leftTotal),rightTotal:R.fmt(report.rightTotal),counts,moneyUnit:'输入金额单位，保留两位小数；JSON 整数金额单位为分',differenceRule:'左表减右表；重复或异常不计算差额'};
 fs.writeFileSync(path.join(out,'摘要.json'),JSON.stringify(summary,null,2),{flag:'wx'});
 const outputs=['差异表.csv','完整核对表.csv','核对记录.json','摘要.json'].map(file=>({file,sha256:fingerprint(path.join(out,file))}));
 const manifest={skill:'finance-reconcile',version:VERSION,runAt:new Date().toISOString(),inputs,fieldMapping:{left:{keys:lk.map(i=>left.headers[i]),amount:opts['--left-amount']},right:{keys:rk.map(i=>right.headers[i]),amount:opts['--right-amount']}},rules:{id:'ordered-tuple-trim-case-sensitive',amount:'integer-cents-exact',duplicates:'retain-all-no-aggregation',delta:'left-minus-right'},coverage,outputs};
 fs.writeFileSync(path.join(out,'运行清单.json'),JSON.stringify(manifest,null,2),{flag:'wx'});
 console.log(JSON.stringify(summary,null,2));
}
try{main();}catch(e){console.error('核对失败：'+e.message);process.exitCode=1;}
