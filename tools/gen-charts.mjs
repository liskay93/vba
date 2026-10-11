// ChartFormatter.bas 의 설정값으로 미리보기 차트 SVG 를 생성합니다.
//   node tools/gen-charts.mjs && node tools/build-preview.mjs
// 차트 SVG 생성기 — ChartFormatter.bas 의 설정값을 그대로 좌표로 옮김
const PT = {W:311, H:215};                 // CHART_WIDTH / CHART_HEIGHT
const F = {base:8.5, legend:9};            // FONT_SIZE_BASE / FONT_SIZE_LEGEND
const AXW = 0.25, GAPW = 80, OVL = -5;     // AXIS_WEIGHT / GAP_WIDTH / OVERLAP_CLUST
const LW = 1.5, MK = 5, FILLGAP = 1.5;     // LINE_WEIGHT / MARKER_SIZE / FILL_GAP_WEIGHT
const INK = "#000000", SURF = "#FFFFFF";   // C_TEXT / C_SURFACE
const FONT = "font-family=\"'Malgun Gothic','맑은 고딕','Noto Sans KR',sans-serif\"";

// 범주형: 1~4번 전 차트 안전, 5번은 막대/꺾은선 전용
const CAT = ["#0B45C4","#5C8FE0","#2E3338","#62686F","#00246B"];
const SEQ = ["#00246B","#0B45C4","#4A7FD9","#86B6EF"];
const OLD = ["#0039B6","#BDD2FF","#DDE9F7","#7F7F7F","#D8D8D8","#00C800"];
// 색약 시야 색과 레이블 글자색은 손으로 옮기지 않고 여기서 계산합니다.
// (Machado 2009, severity 1.0 - cvd-simulate.mjs 및 검증기와 동일)
const _M={protan:[[0.152286,1.052583,-0.204868],[0.114503,0.786281,0.099216],[-0.003882,-0.048116,1.051998]]};
const _lin=h=>h.replace(/^#/,"").match(/../g).map(x=>parseInt(x,16)/255)
  .map(c=>c<=0.04045?c/12.92:((c+0.055)/1.055)**2.4);
const _s=c=>{c=Math.max(0,Math.min(1,c));return c<=0.0031308?12.92*c:1.055*c**(1/2.4)-0.055};
const simulate=(h,k)=>{const[r,g,b]=_lin(h),M=_M[k],q=x=>Math.max(0,Math.min(1,x));
 return "#"+[q(M[0][0]*r+M[0][1]*g+M[0][2]*b),q(M[1][0]*r+M[1][1]*g+M[1][2]*b),q(M[2][0]*r+M[2][1]*g+M[2][2]*b)]
   .map(c=>Math.round(_s(c)*255).toString(16).padStart(2,"0")).join("").toUpperCase()};
// ChartFormatter.bas 의 LabelInk 와 같은 규칙: 조각 밝기에 따라 흰/검정
const labelInk=h=>{const[r,g,b]=_lin(h);const lum=0.2126*r+0.7152*g+0.0722*b;
  return (1.05/(lum+0.05)) >= ((lum+0.05)/0.05) ? SURF : INK};
const CATp = CAT.map(h=>simulate(h,"protan"));
const OLDp = OLD.map(h=>simulate(h,"protan"));

// 그림 영역
const P = {l:26, r:306, t:26, b:184};
const LEGEND_BOXES = [];
const esc = s => s.replace(/&/g,"&amp;").replace(/</g,"&lt;");

// 글자폭 근사: 맑은 고딕 기준 한글 1.0em, 라틴/숫자 0.52em, 공백 0.25em
function textW(s, size){
  let w = 0;
  for(const ch of s){
    if(ch === " ") w += 0.25;
    else if(/[\u1100-\u11FF\u3130-\u318F\uAC00-\uD7AF]/.test(ch)) w += 1.0;
    else w += 0.52;
  }
  return w*size;
}
function legend(names, colors, kind){
  // 범례: 위쪽, 가로 한 줄, 가운데 정렬. 테두리·채우기 없음.
  const sw = 7, gapSL = 3, gapItem = 11;
  const widths = names.map(n => sw + gapSL + textW(n, F.legend));
  const total = widths.reduce((a,b)=>a+b,0) + gapItem*(names.length-1);
  if(total > PT.W - 8) throw new Error(`범례가 차트 폭을 넘침: ${total.toFixed(1)} > ${PT.W-8}`);
  LEGEND_BOXES.length = 0;
  let x = (PT.W - total)/2, out = "";
  names.forEach((n,i)=>{
    LEGEND_BOXES.push([x, x+widths[i]]);
    const cy = 10;
    if(kind==="line"){
      out += `<line x1="${f(x)}" y1="${cy}" x2="${f(x+sw)}" y2="${cy}" stroke="${colors[i]}" stroke-width="${LW}"/>`
           + `<circle cx="${f(x+sw/2)}" cy="${cy}" r="${MK/2}" fill="${colors[i]}" stroke="${SURF}" stroke-width="0.75"/>`;
    } else {
      out += `<rect x="${f(x)}" y="${cy-3.5}" width="${sw}" height="${sw}" fill="${colors[i]}" stroke="${SURF}" stroke-width="0.5"/>`;
    }
    out += `<text x="${f(x+sw+gapSL)}" y="${cy+3.2}" ${FONT} font-size="${F.legend}" fill="${INK}">${esc(n)}</text>`;
    x += widths[i] + gapItem;
  });
  return out;
}
const f = n => Math.round(n*100)/100;

function axes(cats, vmax, vstep){
  let out = "";
  // 값 축: 선 제거(SHOW_GRIDLINES=False 이므로 그리드도 없음), 눈금 안쪽, 라벨만
  for(let v=0; v<=vmax; v+=vstep){
    const y = P.b - (v/vmax)*(P.b-P.t);
    out += `<line x1="${P.l}" y1="${f(y)}" x2="${P.l+2}" y2="${f(y)}" stroke="${INK}" stroke-width="${AXW}"/>`
         + `<text x="${P.l-3}" y="${f(y+3)}" ${FONT} font-size="${F.base}" fill="${INK}" text-anchor="end">${v}</text>`;
  }
  out += `<line x1="${P.l}" y1="${P.t}" x2="${P.l}" y2="${P.b}" stroke="${INK}" stroke-width="${AXW}"/>`;
  // 항목 축: 검정 0.25pt + 눈금 안쪽
  out += `<line x1="${P.l}" y1="${P.b}" x2="${P.r}" y2="${P.b}" stroke="${INK}" stroke-width="${AXW}"/>`;
  const slot = (P.r-P.l)/cats.length;
  cats.forEach((c,i)=>{
    const xc = P.l + slot*(i+0.5);
    out += `<line x1="${f(P.l+slot*(i+1))}" y1="${P.b}" x2="${f(P.l+slot*(i+1))}" y2="${P.b-2}" stroke="${INK}" stroke-width="${AXW}"/>`
         + `<text x="${f(xc)}" y="${P.b+10}" ${FONT} font-size="${F.base}" fill="${INK}" text-anchor="middle">${esc(c)}</text>`;
  });
  return out;
}

function barChart(cats, series, colors, vmax, vstep, stacked){
  const slot = (P.r-P.l)/cats.length, N = series.length;
  let out = legend(series.map(s=>s.name), colors, "bar") + axes(cats, vmax, vstep);
  const h = v => (v/vmax)*(P.b-P.t);
  if(stacked){
    // 누적: Overlap = 100
    const bw = slot/(1 + GAPW/100);
    cats.forEach((_,ci)=>{
      let acc = 0;
      const x = P.l + slot*ci + (slot-bw)/2;
      series.forEach((s,si)=>{
        const y = P.b - h(acc+s.v[ci]);
        out += `<rect x="${f(x)}" y="${f(y)}" width="${f(bw)}" height="${f(h(s.v[ci]))}" fill="${colors[si]}" stroke="${SURF}" stroke-width="${FILLGAP}"/>`;
        acc += s.v[ci];
      });
    });
  } else {
    // 묶은 막대: slot = barW*(N - (N-1)*OVL/100) + barW*GAPW/100
    const bw = slot/((N - (N-1)*OVL/100) + GAPW/100);
    const gap = bw*GAPW/100, step = bw*(1 - OVL/100);
    cats.forEach((_,ci)=>{
      const x0 = P.l + slot*ci + gap/2;
      series.forEach((s,si)=>{
        const y = P.b - h(s.v[ci]);
        out += `<rect x="${f(x0+si*step)}" y="${f(y)}" width="${f(bw)}" height="${f(h(s.v[ci]))}" fill="${colors[si]}" stroke="${SURF}" stroke-width="${FILLGAP}"/>`;
      });
    });
  }
  return out;
}

function lineChart(cats, series, colors, vmax, vstep){
  const slot = (P.r-P.l)/cats.length;
  let out = legend(series.map(s=>s.name), colors, "line") + axes(cats, vmax, vstep);
  const xf = i => P.l + slot*(i+0.5);
  const yf = v => P.b - (v/vmax)*(P.b-P.t);
  series.forEach((s,si)=>{
    const pts = s.v.map((v,i)=>`${f(xf(i))},${f(yf(v))}`).join(" ");
    // Smooth = False -> 직선
    out += `<polyline points="${pts}" fill="none" stroke="${colors[si]}" stroke-width="${LW}" stroke-linejoin="round" stroke-linecap="round"/>`;
    s.v.forEach((v,i)=>{ out += `<circle cx="${f(xf(i))}" cy="${f(yf(v))}" r="${MK/2}" fill="${colors[si]}" stroke="${SURF}" stroke-width="0.75"/>`; });
  });
  return out;
}

function donut(slices, colors, opt={}){
  const cx = PT.W/2, cy = opt.cy ?? 108, R = opt.R ?? 68, r = R*0.5; // DoughnutHoleSize = 50
  const tot = slices.reduce((a,s)=>a+s.v,0);
  let a0 = -Math.PI/2, out = "";
  const lbl = [];
  slices.forEach((s,i)=>{
    const a1 = a0 + 2*Math.PI*s.v/tot, big = (a1-a0) > Math.PI ? 1 : 0;
    const p = (rad,a)=>`${f(cx+rad*Math.cos(a))},${f(cy+rad*Math.sin(a))}`;
    out += `<path d="M ${p(R,a0)} A ${R} ${R} 0 ${big} 1 ${p(R,a1)} L ${p(r,a1)} A ${r} ${r} 0 ${big} 0 ${p(r,a0)} Z" `
         + `fill="${colors[i]}" stroke="${SURF}" stroke-width="${FILLGAP}"/>`;
    if(opt.labels){
      const am = (a0+a1)/2, rm = (R+r)/2;
      const ink = labelInk(colors[i]);   // 연한 조각 위의 흰 글씨는 읽히지 않음
      lbl.push(`<text x="${f(cx+rm*Math.cos(am))}" y="${f(cy+rm*Math.sin(am)-1)}" ${FONT} font-size="${F.base}" font-weight="bold" fill="${ink}" text-anchor="middle">${esc(s.name)}</text>`
             + `<text x="${f(cx+rm*Math.cos(am))}" y="${f(cy+rm*Math.sin(am)+8)}" ${FONT} font-size="${F.base}" font-weight="bold" fill="${ink}" text-anchor="middle">${Math.round(s.v/tot*100)}%</text>`);
    }
    a0 = a1;
  });
  // 레이블이 있으면 범례 생략 (PIE_DATA_LABELS = True -> HasLegend = False)
  const leg = opt.labels ? "" : legend(slices.map(s=>s.name), colors, "bar");
  return leg + out + lbl.join("");
}

// ── 데이터 (ESG 리포트 예시) ─────────────────────────────────────────────
const YRS = ["2021","2022","2023","2024","2025"];
const BAR = [
  {name:"당사", v:[18.4,17.1,15.8,14.2,12.9]},
  {name:"동종업계 평균", v:[21.0,20.4,19.9,19.1,18.6]},
  {name:"업종 상위 25%", v:[15.2,14.6,14.0,13.3,12.7]},
];
const LINE = [
  {name:"재생에너지", v:[8,12,19,27,38,46]},
  {name:"에너지 집약도", v:[42,39,36,33,30,28]},
  {name:"폐기물 재활용", v:[61,64,66,70,73,76]},
  {name:"용수 재이용", v:[15,18,22,25,29,34]},
];
const LYRS = ["2020","2021","2022","2023","2024","2025"];
const DONUT = [{name:"전력",v:42},{name:"도시가스",v:27},{name:"스팀",v:19},{name:"기타",v:12}];
const STACK = [
  {name:"Scope 1", v:[6.1,5.8,5.4,5.0,4.6]},
  {name:"Scope 2", v:[7.3,6.9,6.4,5.8,5.3]},
  {name:"Scope 3", v:[4.0,3.8,3.5,3.2,2.9]},
];

function svg(inner, h){
  return `<svg class="chart" viewBox="0 0 ${PT.W} ${h ?? PT.H}" role="img" preserveAspectRatio="xMidYMid meet">`
       + `<rect x="0" y="0" width="${PT.W}" height="${h ?? PT.H}" fill="${SURF}"/>${inner}</svg>`;
}
const out = {
  bar:     svg(barChart(YRS, BAR, CAT, 25, 5, false)),
  line:    svg(lineChart(LYRS, LINE, CAT, 80, 20)),
  donut:   svg(donut(DONUT, CAT, {labels:true})),
  barOld:  svg(barChart(YRS, BAR, OLD, 25, 5, false)),
  stack:   svg(barChart(YRS, STACK, SEQ, 20, 5, true)),
  dNewN:   svg(donut(DONUT, CAT,  {cy:76, R:62}), 152),
  dNewP:   svg(donut(DONUT, CATp, {cy:76, R:62}), 152),
  dOldN:   svg(donut(DONUT, OLD,  {cy:76, R:62}), 152),
  dOldP:   svg(donut(DONUT, OLDp, {cy:76, R:62}), 152),
};
// 범례 항목 겹침 검증
function assertNoOverlap(tag){
  for(let i=1;i<LEGEND_BOXES.length;i++){
    const prev=LEGEND_BOXES[i-1], cur=LEGEND_BOXES[i];
    if(cur[0] < prev[1]) throw new Error(`${tag}: 범례 ${i}번이 ${i-1}번과 겹침 (${prev[1].toFixed(1)} > ${cur[0].toFixed(1)})`);
  }
  const gaps=LEGEND_BOXES.slice(1).map((c,i)=>c[0]-LEGEND_BOXES[i][1]);
  console.log(`  ${tag}: 항목 ${LEGEND_BOXES.length}개, 최소 간격 ${Math.min(...gaps).toFixed(1)}pt, 전체 폭 ${(LEGEND_BOXES.at(-1)[1]-LEGEND_BOXES[0][0]).toFixed(1)}pt / ${PT.W}pt`);
}
console.log("범례 겹침 검증:");
barChart(YRS, BAR, CAT, 25, 5, false); assertNoOverlap("막대");
lineChart(LYRS, LINE, CAT, 80, 20);    assertNoOverlap("꺾은선");
barChart(YRS, STACK, SEQ, 20, 5, true);assertNoOverlap("누적");

import("fs").then(async fs=>{
  const dir = new URL(".", import.meta.url).pathname;
  fs.writeFileSync(dir + "charts.json", JSON.stringify(out));
  console.log("생성 완료:", Object.keys(out).join(", "));
});

// 기하 검증
const slot=(P.r-P.l)/5, N=3, bw=slot/((N-(N-1)*OVL/100)+GAPW/100);
console.log(`\n막대 기하 검증 (GapWidth=${GAPW}, Overlap=${OVL}):`);
console.log(`  슬롯폭 ${f(slot)}pt, 막대폭 ${f(bw)}pt`);
console.log(`  클러스터+간격 = ${f(bw*(N-(N-1)*OVL/100) + bw*GAPW/100)}pt  (슬롯폭과 같아야 함: ${f(slot)})`);
console.log(`  도넛 구멍 = 50% (R=68 -> r=34)`);
