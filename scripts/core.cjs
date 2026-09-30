(function(root){
'use strict';
const MAX=2*1024*1024;
function parse(text){
 if(text.length>MAX||new TextEncoder().encode(text).length>MAX)throw Error('每份输入请控制在 2 MB 以内。');
 text=text.replace(/^\uFEFF/,'');
 let quoted=false,comma=0,tab=0;
 for(let i=0;i<text.length;i++){let c=text[i];if(c==='"'){if(quoted&&text[i+1]==='"')i++;else quoted=!quoted;}else if(!quoted){if(c==='\n'||c==='\r')break;if(c===',')comma++;if(c==='\t')tab++;}}
 const sep=tab>comma?'\t':',';const rows=[];let row=[],field='',state=0;
 function push(){row.push(field);field='';state=0;}
 function line(){push();if(row.some(x=>x.trim()!==''))rows.push(row);row=[];}
 for(let i=0;i<text.length;i++){
  const c=text[i];
  if(state===1){if(c==='"'){if(text[i+1]==='"'){field+='"';i++;}else state=2;}else field+=c;continue;}
  if(c===sep){push();continue;}
  if(c==='\n'||c==='\r'){if(c==='\r'&&text[i+1]==='\n')i++;line();continue;}
  if(state===2){if(c===' '||c==='\t'&&sep!== '\t')continue;throw Error('引号字段结束后存在多余字符。');}
  if(c==='"'){if(field!=='')throw Error('引号应位于字段开头。');state=1;}else field+=c;
 }
 if(state===1)throw Error('CSV 引号未闭合。');
 if(field!==''||row.length||state===2)line();
 if(!rows.length)throw Error('请提供表头和明细记录。');
 const headers=rows.shift().map(x=>x.trim());
 if(headers.some(x=>!x)||new Set(headers).size!==headers.length)throw Error('表头不能为空或重复。');
 if(headers.length>50||rows.length>5000)throw Error('每份数据最多 50 列、5000 行。');
 rows.forEach((r,i)=>{if(r.length!==headers.length)throw Error(`第 ${i+2} 条表格记录的列数与表头不一致。`);});
 if(!rows.length)throw Error('表头下面还没有明细。');
 return {headers,rows};
}
function money(value){
 let s=String(value).trim();if(!s||s.length>100)return null;
 if(/^\([\d,.]+\)$/.test(s))s='-'+s.slice(1,-1);
 if(!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?$/.test(s))return null;
 s=s.replace(/,/g,'');const negative=s[0]==='-';s=s.replace(/^[+-]/,'');
 const [a,b='']=s.split('.');const n=BigInt(a)*100n+BigInt(b.padEnd(2,'0'));return negative?-n:n;
}
function fmt(n){if(n===null)return '不可计算';const neg=n<0n;if(neg)n=-n;return (neg?'-':'')+(n/100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g,',')+'.'+(n%100n).toString().padStart(2,'0');}
function reconcile(left,right,lk,la,rk,ra){
 function prep(data,key,amount){
  if(!Number.isInteger(key)||!Number.isInteger(amount)||key<0||amount<0||key>=data.headers.length||amount>=data.headers.length||key===amount)throw Error('编号列与金额列需要选择不同的有效列。');
  const groups=new Map(),invalid=[];let total=0n,bad=0;
  data.rows.forEach((r,i)=>{const id=r[key].trim(),n=money(r[amount]);const item={line:data.sourceLines?.[i]??i+2,id,raw:r[amount],amount:n};if(n===null)bad++;else total+=n;if(!id)invalid.push(item);else {if(!groups.has(id))groups.set(id,[]);groups.get(id).push(item);}});
  return {groups,invalid,total:bad?null:total,bad};
 }
 const l=prep(left,lk,la),r=prep(right,rk,ra),result=[];
 for(const id of new Set([...l.groups.keys(),...r.groups.keys()])){
  const a=l.groups.get(id)||[],b=r.groups.get(id)||[];let status,delta=null;
  if(a.length>1||b.length>1)status='重复编号';
  else if([...a,...b].some(x=>x.amount===null))status='无效金额';
  else if(!a.length)status='仅右表';else if(!b.length)status='仅左表';
  else {delta=a[0].amount-b[0].amount;status=delta===0n?'一致':'金额差异';}
  result.push({id,status,left:a,right:b,delta});
 }
 for(const [side,items] of [['left',l.invalid],['right',r.invalid]])for(const x of items)result.push({id:'（空编号）',status:x.amount===null?'空编号及无效金额':'空编号',left:side==='left'?[x]:[],right:side==='right'?[x]:[],delta:null});
 return {result,leftTotal:l.total,rightTotal:r.total,leftBad:l.bad,rightBad:r.bad,leftRows:left.rows.length,rightRows:right.rows.length};
}
function csvCell(value){let s=String(value);if(/^[\s]*[=+\-@]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';}
function exportCSV(rows){const data=[['编号','状态','左表记录序号','右表记录序号','左表原金额','右表原金额','左减右差额'],...rows.map(x=>[x.id,x.status,x.left.map(r=>r.line).join(' / '),x.right.map(r=>r.line).join(' / '),x.left.map(r=>r.raw).join(' / '),x.right.map(r=>r.raw).join(' / '),x.delta===null?'待核实':fmt(x.delta)])];return '\uFEFF'+data.map(r=>r.map(csvCell).join(',')).join('\r\n');}
const api={parse,money,fmt,reconcile,exportCSV};if(typeof module!=='undefined'&&module.exports)module.exports=api;else root.Reconcile=api;
})(typeof window!=='undefined'?window:globalThis);
