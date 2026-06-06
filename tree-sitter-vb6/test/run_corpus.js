const fs=require("fs"),cp=require("child_process"),path=require("path");
const exe=path.join(__dirname,"vb6parse.exe");
const tmp=path.join(__dirname,"_t.bas");
const norm=s=>s.replace(/\r/g,"").replace(/\s+/g," ").replace(/\(\s+/g,"(").replace(/\s+\)/g,")").trim();
// take a balanced-paren s-expression starting at first '('
function sexp(s){const i=s.indexOf("(");if(i<0)return"";let d=0;for(let j=i;j<s.length;j++){if(s[j]==="(")d++;else if(s[j]===")"){d--;if(d===0)return s.slice(i,j+1);}}return s.slice(i);}
let pass=0,fail=0,fails=[];
for(const file of fs.readdirSync(path.join(__dirname,"corpus")).filter(f=>f.endsWith(".txt"))){
  const txt=fs.readFileSync(path.join(__dirname,"corpus",file),"utf8").replace(/\r/g,"");
  const re=/(?:^|\n)={3,}\n(.+?)\n={3,}\n([\s\S]*?)\n---\n([\s\S]*?)(?=\n={3,}\n|$)/g;
  let m;
  while((m=re.exec(txt))){
    const name=m[1].trim(), input=m[2].replace(/^\n+|\n+$/g,""), exp=sexp(m[3]);
    if(!exp)continue;
    fs.writeFileSync(tmp,input+"\n");
    let got;
    try{got=cp.execSync(`"${exe}" "${tmp}"`).toString();}catch(e){got="<<crash>>";}
    if(norm(sexp(got))===norm(exp)){pass++;}
    else{fail++;fails.push({file,name,got:norm(sexp(got)),exp:norm(exp)});}
  }
}
fs.unlinkSync(tmp);
console.log(`\nCORPUS RESULT: ${pass} passed, ${fail} failed\n`);
for(const f of fails){console.log(`FAIL [${f.file}] ${f.name}`);console.log("  got: "+f.got);console.log("  exp: "+f.exp+"\n");}
process.exit(fail?1:0);
