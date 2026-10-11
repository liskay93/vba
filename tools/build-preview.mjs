// charts.json 을 읽어 preview.html 을 조립합니다. gen-charts.mjs 를 먼저 실행하세요.
import fs from "fs";
const DIR = new URL(".", import.meta.url).pathname;
const C = JSON.parse(fs.readFileSync(DIR + "charts.json","utf8"));
const H = `<title>ESG 차트 서식 교정쇄</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&family=Noto+Sans+KR:wght@400;700&display=swap">
<style>
/* 레이아웃: 단일 컬럼 교정쇄 — 헤어라인으로 구획하고, 차트는 실제 리포트 크기의 흰 지면에 올린다 */
:root{
  --paper:#FFFFFF;            /* 차트 지면: 모듈이 C_SURFACE 로 고정하므로 테마와 무관 */
  --paper-ink:#000000;        /* 차트 글자: C_TEXT */
  --bg:#F6F7FA; --card:#FFFFFF; --ink:#14171F; --muted:#59606E;
  --rule:#E0E4EC; --rule-soft:#EDF0F5; --accent:#0B45C4; --warn:#8A4B00;
  --font-ui:'IBM Plex Sans KR',system-ui,sans-serif;
  --font-data:'IBM Plex Mono',ui-monospace,monospace;
}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
  --bg:#101219; --card:#171A22; --ink:#E7EAF0; --muted:#949BAA;
  --rule:#272C37; --rule-soft:#1F232C; --accent:#7CA4F2; --warn:#D9A24B; color-scheme:dark;}}
:root[data-theme="dark"]{
  --bg:#101219; --card:#171A22; --ink:#E7EAF0; --muted:#949BAA;
  --rule:#272C37; --rule-soft:#1F232C; --accent:#7CA4F2; --warn:#D9A24B; color-scheme:dark;}

body{background:var(--bg);color:var(--ink);font-family:var(--font-ui);line-height:1.65;
 -webkit-font-smoothing:antialiased}
.wrap{max-width:920px;margin:0 auto;padding-inline:20px;padding-block:44px 72px}
h1{font-size:clamp(1.55rem,4vw,2.1rem);font-weight:600;letter-spacing:-.02em;margin:0;text-wrap:balance}
.lede{color:var(--muted);max-width:60ch;margin:.7rem 0 0}
.eyebrow{font-family:var(--font-data);font-size:.68rem;font-weight:500;letter-spacing:.14em;
 text-transform:uppercase;color:var(--accent);margin:0 0 .35rem}
.note{display:flex;gap:.6rem;align-items:flex-start;margin-top:1.4rem;padding:.75rem .9rem;
 border:1px solid var(--rule);border-left:2px solid var(--warn);border-radius:3px;
 background:var(--card);font-size:.84rem;color:var(--muted);max-width:66ch}
.note b{color:var(--ink);font-weight:600}
section{margin-top:3.2rem;padding-top:2rem;border-top:1px solid var(--rule)}
section>h2{font-size:1.12rem;font-weight:600;margin:0 0 .3rem;letter-spacing:-.01em}
section>p.sub{color:var(--muted);font-size:.9rem;margin:0 0 1.5rem;max-width:62ch}

/* 교정쇄 지면 */
.proof{background:var(--card);border:1px solid var(--rule);border-radius:4px;overflow:hidden}
.proof-hd{display:flex;justify-content:space-between;align-items:baseline;gap:1rem;
 padding:.6rem .85rem;border-bottom:1px solid var(--rule-soft);flex-wrap:wrap}
.proof-hd .t{font-size:.86rem;font-weight:600}
.proof-hd .d{font-family:var(--font-data);font-size:.7rem;color:var(--muted);
 font-variant-numeric:tabular-nums}
.sheet{background:var(--paper);padding:14px;display:flex;justify-content:center}
.chart{width:100%;max-width:415px;height:auto;display:block}
.cap{padding:.65rem .85rem;border-top:1px solid var(--rule-soft);font-size:.79rem;color:var(--muted)}
.cap code{font-family:var(--font-data);font-size:.95em;color:var(--ink);background:var(--rule-soft);
 padding:.08em .32em;border-radius:2px}
.stack{display:flex;flex-direction:column;gap:1.4rem}
.duo{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,290px),1fr));gap:1.1rem}
.quad{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,215px),1fr));gap:.9rem}
.quad .chart{max-width:300px}
.quad .sheet{padding:8px}

/* 팔레트 띠 */
.sw{display:flex;flex-wrap:wrap;gap:.45rem;margin:.1rem 0 0;padding:0;list-style:none}
.sw li{display:flex;align-items:center;gap:.4rem;font-family:var(--font-data);font-size:.72rem;
 color:var(--muted);font-variant-numeric:tabular-nums;min-width:0}
.sw i{width:13px;height:13px;border-radius:2px;flex:none;box-shadow:inset 0 0 0 1px rgba(0,0,0,.14)}
.bad{color:var(--warn);font-weight:500}

table{width:100%;border-collapse:collapse;font-size:.83rem}
.tw{overflow-x:auto}
th,td{text-align:left;padding:.44rem .7rem;border-bottom:1px solid var(--rule-soft);vertical-align:top}
th{font-family:var(--font-data);font-size:.68rem;font-weight:500;letter-spacing:.09em;
 text-transform:uppercase;color:var(--muted);white-space:nowrap}
td.n{font-family:var(--font-data);font-variant-numeric:tabular-nums;white-space:nowrap}
ul.plain{margin:.2rem 0 0;padding-left:1.1rem;color:var(--muted);font-size:.87rem}
ul.plain li{margin-bottom:.45rem}
ul.plain b{color:var(--ink);font-weight:600}
code{font-family:var(--font-data)}
</style>

<div class="wrap">
<p class="eyebrow">ChartFormatter / 적용 결과</p>
<h1>ESG 차트 서식 교정쇄</h1>
<p class="lede">차트를 선택하고 <code>FormatSelectedChart</code> 를 실행했을 때 나오는 모양입니다.
모듈의 설정 상수를 그대로 좌표로 옮겨 그렸습니다 — 크기 311 × 215 pt, 본문 8.5 pt,
축 검정 0.25 pt, 막대 간격 너비 80, 겹치기 −5, 도넛 구멍 50 %.</p>

<div class="note"><span>※</span><span><b>엑셀 스크린샷이 아니라 SVG 재현입니다.</b>
글꼴은 이 화면에 맑은 고딕이 없으면 Noto Sans KR 로 대체되고, 범례용
KoPubWorld돋움체 Medium 은 대체 글꼴로 보입니다. 엑셀의 자동 축 눈금 선택과
레이블 충돌 회피는 재현하지 않았습니다.</span></div>

<section>
<h2>세 가지 차트 종류</h2>
<p class="sub">각 지면은 실제 출력 크기(311 × 215 pt)입니다. 범례는 위쪽, 그리드는 제거,
값 축은 눈금만 안쪽으로 남습니다.</p>
<div class="stack">

  <div class="proof">
    <div class="proof-hd"><span class="t">묶은 막대 — 온실가스 배출 집약도</span>
      <span class="d">계열 3 · 슬롯 56 pt · 막대 14.36 pt</span></div>
    <div class="sheet">${C.bar}</div>
    <div class="cap">계열이 2개 이상이므로 범례가 위쪽에 켜집니다. 막대 경계는
      <code>#FFFFFF</code> 1.5 pt 선으로, 붙어 있는 막대를 떼어 놓습니다.
      단위 tCO₂e/억원.</div>
  </div>

  <div class="proof">
    <div class="proof-hd"><span class="t">꺾은선 — 환경 지표 추이</span>
      <span class="d">계열 4 · 선 1.5 pt · 마커 5 pt</span></div>
    <div class="sheet">${C.line}</div>
    <div class="cap">곡선 보정은 해제(<code>Smooth = False</code>)되어 직선으로 이어집니다.
      이미 마커가 있는 계열만 5 pt + 흰 테두리로 손보고, 없던 마커는 만들지 않습니다. 단위 %.</div>
  </div>

  <div class="proof">
    <div class="proof-hd"><span class="t">도넛 — 에너지 사용 구성</span>
      <span class="d">조각 4 · 구멍 50 % · 레이블 가운데</span></div>
    <div class="sheet">${C.donut}</div>
    <div class="cap">조각별로 색이 배정되고, 레이블이 항목명과 비중을 말해주므로
      범례는 켜지지 않습니다(<code>PIE_DATA_LABELS = True</code>). 글자는 흰색 굵게.</div>
  </div>

</div>
</section>

<section>
<h2>팔레트를 바꾼 이유</h2>
<p class="sub">같은 데이터, 같은 서식, 색만 다릅니다. 기존 팔레트의 2·3번은 파랑 한 색의
진하기 단계여서, 서로 다른 대상을 구분하는 용도로 쓰면 흰 바탕에서 사라집니다.</p>
<div class="duo">
  <div class="proof">
    <div class="proof-hd"><span class="t">기존</span><span class="d bad">6색 중 4색이 3:1 미만</span></div>
    <div class="sheet">${C.barOld}</div>
    <div class="cap"><ul class="sw">
      <li><i style="background:#0039B6"></i>#0039B6 9.29</li>
      <li><i style="background:#BDD2FF"></i><span class="bad">#BDD2FF 1.52</span></li>
      <li><i style="background:#DDE9F7"></i><span class="bad">#DDE9F7 1.23</span></li>
    </ul></div>
  </div>
  <div class="proof">
    <div class="proof-hd"><span class="t">권장</span><span class="d">6색 전부 3:1 이상</span></div>
    <div class="sheet">${C.bar}</div>
    <div class="cap"><ul class="sw">
      <li><i style="background:#0B45C4"></i>#0B45C4 7.89</li>
      <li><i style="background:#E8730C"></i>#E8730C 3.05</li>
      <li><i style="background:#00897B"></i>#00897B 4.32</li>
    </ul></div>
  </div>
</div>
</section>

<section>
<h2>색약 시야</h2>
<p class="sub">도넛은 모든 조각이 한눈에 같이 보이므로 구분 기준이 가장 엄격합니다.
적색맹 시야는 Machado 2009 변환(severity 1.0)으로 계산했습니다.</p>
<div class="quad">
  <div class="proof"><div class="proof-hd"><span class="t">기존 · 정상</span></div>
    <div class="sheet">${C.dOldN}</div></div>
  <div class="proof"><div class="proof-hd"><span class="t">기존 · 적색맹</span>
    <span class="d bad">최저 ΔE 7.4</span></div>
    <div class="sheet">${C.dOldP}</div></div>
  <div class="proof"><div class="proof-hd"><span class="t">권장 · 정상</span></div>
    <div class="sheet">${C.dNewN}</div></div>
  <div class="proof"><div class="proof-hd"><span class="t">권장 · 적색맹</span>
    <span class="d">최저 ΔE 8.1</span></div>
    <div class="sheet">${C.dNewP}</div></div>
</div>
<div class="note"><span>※</span><span>권장 팔레트도 여유가 크지는 않습니다. 적색맹 시야에서
주황과 황토가 ΔE 8.1 로 기준선(8.0)을 겨우 넘습니다. <b>그래서 도넛은 4조각까지입니다</b> —
5조각이 되면 자주(<code>#7B1FA2</code>)가 브랜드 블루 쪽으로 수렴합니다. 조각이 5개를 넘으면
매크로가 경고를 띄웁니다.</span></div>
</section>

<section>
<h2>순차 모드</h2>
<p class="sub"><code>PALETTE_MODE = "SEQUENTIAL"</code> 로 바꾸면 파랑 4단계를 씁니다.
누적 막대나 등급처럼 순서가 있는 구간에 쓰는 모드입니다.</p>
<div class="proof">
  <div class="proof-hd"><span class="t">누적 막대 — Scope 1·2·3</span>
    <span class="d">Overlap 100 고정 · 4단계 램프</span></div>
  <div class="sheet">${C.stack}</div>
  <div class="cap">누적 차트는 <code>Overlap</code> 을 100으로 강제합니다. 다른 값이면 누적이
    시각적으로 깨집니다. 램프 색: <ul class="sw" style="margin-top:.4rem">
      <li><i style="background:#00246B"></i>#00246B 14.34</li>
      <li><i style="background:#0B45C4"></i>#0B45C4 7.89</li>
      <li><i style="background:#4A7FD9"></i>#4A7FD9 3.94</li>
      <li><i style="background:#86B6EF"></i>#86B6EF 2.11</li>
    </ul></div>
</div>
</section>

<section>
<h2>적용되는 서식값</h2>
<div class="tw"><table>
<thead><tr><th>대상</th><th>값</th><th>상수</th></tr></thead>
<tbody>
<tr><td>본문 글꼴</td><td class="n">맑은 고딕 8.5 pt</td><td class="n">FONT_BODY</td></tr>
<tr><td>범례 글꼴</td><td class="n">KoPubWorld돋움체 Medium 9 pt</td><td class="n">FONT_LEGEND</td></tr>
<tr><td>모든 차트 글자</td><td class="n">#000000</td><td class="n">C_TEXT</td></tr>
<tr><td>축 선</td><td class="n">#000000 · 0.25 pt · 눈금 안쪽</td><td class="n">AXIS_WEIGHT</td></tr>
<tr><td>그리드</td><td class="n">제거</td><td class="n">SHOW_GRIDLINES = False</td></tr>
<tr><td>범례 위치</td><td class="n">위쪽 · 테두리 없음</td><td class="n">xlLegendPositionTop</td></tr>
<tr><td>막대</td><td class="n">간격 80 · 겹치기 −5 · 경계 흰색 1.5 pt</td><td class="n">GAP_WIDTH</td></tr>
<tr><td>꺾은선</td><td class="n">1.5 pt · 곡선 보정 해제</td><td class="n">LINE_WEIGHT</td></tr>
<tr><td>크기</td><td class="n">건드리지 않음 (기본값)</td><td class="n">APPLY_SIZE = False</td></tr>
<tr><td>위치</td><td class="n">저장 후 복원 — 변하지 않음</td><td class="n">—</td></tr>
</tbody></table></div>
<p class="sub" style="margin-top:1.1rem">위 지면들은 <code>APPLY_SIZE = True</code> 로 두고
311 × 215 pt 를 적용한 모습입니다. 기본값은 <code>False</code> 이므로 실제로는 직접 만든
차트의 크기가 그대로 유지됩니다.</p>
</section>

<section>
<h2>재현하지 않은 것</h2>
<ul class="plain">
<li><b>엑셀의 자동 축 눈금.</b> 여기서는 눈금 간격을 직접 정했습니다. 실제로는 엑셀이
데이터 범위에 따라 고르고, 모듈은 <code>VALUE_AXIS_FORMAT</code> 이 빈 값이면 숫자 서식을
건드리지 않습니다.</li>
<li><b>레이블 충돌 회피.</b> 엑셀은 항목 축 레이블이 겹치면 기울이거나 건너뜁니다.
도넛 조각이 좁으면 레이블이 조각 밖으로 밀려나기도 합니다.</li>
<li><b>글꼴 실물.</b> 이 화면에 맑은 고딕과 KoPubWorld돋움체가 설치되어 있지 않으면
대체 글꼴로 보입니다. 자간과 글자 폭이 실제와 다릅니다.</li>
<li><b>엑셀에서의 실행 검증.</b> 리눅스 환경이라 매크로를 직접 돌려보지는 못했습니다.
막대·꺾은선·도넛 하나씩 돌려보시고 다른 점이 있으면 알려주세요.</li>
</ul>
</section>
</div>`;
fs.writeFileSync(DIR + "../preview.html", H);
console.log("preview.html", (H.length/1024).toFixed(1)+"KB");
