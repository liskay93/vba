// 색약 시야 색상 계산. dataviz 검증기(validate_palette.js)와 동일한
// Machado 2009 변환(severity 1.0, linear RGB)을 씁니다.
//   node tools/cvd-simulate.mjs
// validate_palette.js 와 동일한 상수/함수 (Machado 2009, severity 1.0, linear RGB)
const MACHADO={protan:[[0.152286,1.052583,-0.204868],[0.114503,0.786281,0.099216],[-0.003882,-0.048116,1.051998]],
 deutan:[[0.367322,0.860646,-0.227968],[0.280085,0.672501,0.047413],[-0.011820,0.042940,0.968881]]};
const hex2srgb=h=>{h=h.trim().replace(/^#/,"");return[0,2,4].map(i=>parseInt(h.slice(i,i+2),16)/255)};
const s2lin=c=>c<=0.04045?c/12.92:((c+0.055)/1.055)**2.4;
const lin2s=c=>{c=Math.max(0,Math.min(1,c));return c<=0.0031308?12.92*c:1.055*c**(1/2.4)-0.055};
const lin=h=>hex2srgb(h).map(s2lin);
function oklabFromLin([r,g,b]){const l=Math.cbrt(0.4122214708*r+0.5363325363*g+0.0514459929*b),
 m=Math.cbrt(0.2119034982*r+0.6806995451*g+0.1073969566*b),s=Math.cbrt(0.0883024619*r+0.2817188376*g+0.6299787005*b);
 return[0.2104542553*l+0.7936177850*m-0.0040720468*s,1.9779984951*l-2.4285922050*m+0.4505937099*s,
 0.0259040371*l+0.7827717662*m-0.8086757660*s]}
function simulate(h,k){const[r,g,b]=lin(h),M=MACHADO[k],c=x=>Math.max(0,Math.min(1,x));
 return[c(M[0][0]*r+M[0][1]*g+M[0][2]*b),c(M[1][0]*r+M[1][1]*g+M[1][2]*b),c(M[2][0]*r+M[2][1]*g+M[2][2]*b)]}
const toHex=v=>"#"+v.map(c=>Math.round(lin2s(c)*255).toString(16).padStart(2,"0")).join("").toUpperCase();
function dE(a,b,k){const x=oklabFromLin(k?simulate(a,k):lin(a)),y=oklabFromLin(k?simulate(b,k):lin(b));
 return 100*Math.hypot(x[0]-y[0],x[1]-y[1],x[2]-y[2])}

const OLD=["#0039B6","#BDD2FF","#DDE9F7","#7F7F7F","#D8D8D8","#00C800"];
const NEW=["#0B45C4","#E8730C","#00897B","#8A6D00","#7B1FA2","#C2185B"];
const out={};
for(const[name,P]of[["old",OLD],["new",NEW]]){
  out[name]={normal:P,protan:P.map(h=>toHex(simulate(h,"protan"))),deutan:P.map(h=>toHex(simulate(h,"deutan")))};
}
console.log(JSON.stringify(out,null,1));
console.log("\n=== 도넛 4조각 전체쌍 최저 dE (적색맹) ===");
for(const[name,P]of[["기존",OLD],["권장",NEW]]){
  let worst=1e9,pair="";
  const S=P.slice(0,4);
  for(let i=0;i<S.length;i++)for(let j=i+1;j<S.length;j++){
    const d=Math.min(dE(S[i],S[j],"protan"),dE(S[i],S[j],"deutan"));
    if(d<worst){worst=d;pair=S[i]+" vs "+S[j]}}
  console.log(`  ${name}: ${worst.toFixed(1)}  (${pair})`);
}
console.log("\n=== 기존 팔레트 2·3·5번 상호 dE (정상 시야) ===");
console.log("  #BDD2FF vs #DDE9F7:", dE("#BDD2FF","#DDE9F7").toFixed(1));
console.log("  #DDE9F7 vs #D8D8D8:", dE("#DDE9F7","#D8D8D8").toFixed(1));
