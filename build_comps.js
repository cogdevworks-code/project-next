// 엠아이텍 Comps 전용 워크북: 트레이딩 Comps + 트랜잭션 Comps + 함의 밸류에이션 + 출처
const XLSX = require("xlsx");
const wb = XLSX.utils.book_new();

// ===== Sheet 1: Trading Comps (Comparable Companies) =====
const t = [
 ["Comparable Companies Analysis — 상장 메디테크 (스텐트/내시경/GI 인접)","","","","","",""],
 ["기준일 2026-06-24 · EV/지표=TTM, P/E=fwd · 출처 stockanalysis/S&P/Yahoo · 🟡 라이브 재확인 권장","","","","","",""],
 ["",""],
 ["회사","티커","EV/EBITDA(TTM)","EV/EBIT(TTM)","Fwd P/E","EV/Revenue(TTM)","구분/메모"],
 ["Boston Scientific","BSX",14.1,18.8,13.3,3.75,"대형 안정·엠아이텍 인수시도 당사자"],
 ["Medtronic","MDT",12.9,"-","-","-","대형(Q4 FY26 추세)"],
 ["Merit Medical","MMSI",13.6,22.2,16.4,2.82,"중형"],
 ["ConMed","CNMD",8.5,13.0,7.6,1.35,"저평가 중형"],
 ["Olympus","7733.T",8.2,"-","-","-","日 콩글로 디스카운트(혼동주의)"],
 ["Ambu A/S","AMBU.CO",33.0,"-","-","-","고성장 순수 엔도 프리미엄"],
 ["Micro-Tech Nanjing","688029.SS",20.7,18.1,"-",4.5,"★직접 동종(순수 스텐트/엔도)"],
 ["",""],
 ["[요약 통계] ★EV/EBITDA·EV/EBIT 중심","EV/EBITDA","EV/EBIT","","EV/Revenue","",""],
 ["중앙값(median)",{f:'MEDIAN(C5:C11)'},{f:'MEDIAN(D5:D11)'},"",{f:'MEDIAN(F5:F11)'},"",""],
 ["평균(mean)",{f:'AVERAGE(C5:C11)'},{f:'AVERAGE(D5:D11)'},"",{f:'AVERAGE(F5:F11)'},"",""],
 ["동종 순수(Micro·Ambu) EV/EBITDA 평균",{f:'AVERAGE(C10:C11)'},"","","","",""],
 ["",""],
 ["▶ 엠아이텍 (179290) — 비교대상","","","","","",""],
 ["엠아이텍 2025E","179290","7.0 (순현금차감)","-","11.4 (PER)","-","🟡 FnGuide · 단순 시총÷EBITDA는 11~12x"],
 ["메모: 엠아이텍 30%대 OPM·고성장에도 7~12x → 대형(13~14x)·동종 순수(20~33x) 대비 디스카운트.","","","","","",""],
];
const ws1 = XLSX.utils.aoa_to_sheet(t);
ws1["!cols"]=[{wch:22},{wch:12},{wch:18},{wch:15},{wch:12},{wch:18},{wch:36}];
XLSX.utils.book_append_sheet(wb, ws1, "Trading Comps");

// ===== Sheet 2: Transaction Comps (Precedent Transactions) =====
const x = [
 ["Precedent Transactions Analysis — 스텐트/GI/내시경 인접 M&A (정밀본)","","","","","",""],
 ["분모(타겟 재무)는 보도/감사보고서 혼합 🟡 → '보도 기준 EV/Rev 약 X배'로 hedge. 적자 타겟은 이익배수 n/m.","","","","","",""],
 ["",""],
 ["인수자 → 타겟","연도(발표/완료)","딜 규모(EV)","타겟 매출·이익","EV/Revenue","EV/EBIT","신뢰도/메모"],
 ["BSX → Apollo Endosurgery","2022.11발표/2023.4완료","$615M($10/주, 67%프리미엄, 전액현금)","FY22 매출 ~$76M·영업손실(적자)","8.1x","n/m(적자)","🟢공식PR · 엔도루미널+엔도베리아트릭(OverStitch·ESG). 성장형→매출배수로 가격. EBITDA도 n/m"],
 ["Olympus → 태웅메디칼","2023.1발표/2024.1완료/2024.3철회","$370M(선급255.5+earn-out114.5)≈4,900억","FY22 매출 685억·영업익 98억(OPM14%)","7.2x(총)/4.9x(선급)","~50x(총)/~34x(선급)","🟢완료·철회 / 🟡재무 · 최근접 동종(韓 비혈관/GI). EBITDA=D&A미공개 n/a"],
 ["BSX → 엠아이텍 (무산)","2022계약/2023.5해지","지분63.9%@14,500원=2,911억→100%환산~4,556억","FY22 매출 ~562억","~8x(100%기준)","-","🟢딜/🟡멀티플 · 자체 거래 앵커. 해외 경쟁당국 미승인·위약금150억·BSX 9.83% 잔류"],
 ["Merit Medical → Cook(lead mgmt)","2024","$210M","비공개","-","-","🟢 · 인접 carve-out"],
 ["Merit Medical → Biolife","2024","$120M","비공개","-","-","🟢 · 인접 bolt-on(지혈)"],
 ["",""],
 ["[EV/Revenue 레인지]","Apollo 8.1x · 태웅(총) 7.2x · BSX-엠텍 ~8x → 레인지 7.2~8.1x, 중심 ~8x","","","","",""],
 ["",""],
 ["[섹터 트랜잭션 밴드]","","","","","",""],
 ["흑자 medtech EV/EBITDA","10~14x","","","","","FOCUS/M&A Insights 🟡"],
 ["일반 medtech EV/Revenue","4~6x(AI 6~8x+)","","","","","Nelson Advisors(유럽) 🟡"],
 ["★핵심: GI/내시경 전략 바이어는 매출 7~8x 지불(흑자·적자 무관). 흑자(태웅·엠텍)는 EV/Rev 7~8x, 적자성장형(Apollo)도 8x.","","","","","",""],
 ["★리스크 사례: 태웅 철회=인수후 제품데이터가 실사데이터와 불일치(R&W 위반). 듀딜의 데이터 검증 중요성 = 직무 어필 포인트.","","","","","",""],
];
const ws2 = XLSX.utils.aoa_to_sheet(x);
ws2["!cols"]=[{wch:30},{wch:16},{wch:28},{wch:18},{wch:13},{wch:14},{wch:40}];
XLSX.utils.book_append_sheet(wb, ws2, "Transaction Comps");

// ===== Sheet 3: 엠아이텍 함의 밸류에이션 (Comps 대입) =====
const v = [
 ["엠아이텍 함의 밸류에이션 — Comps 대입 (값 바꾸면 자동계산)","","",""],
 ["단위: 억원 · FY2025 실적 기준","","",""],
 ["",""],
 ["[INPUT]","값","","메모"],
 ["FY2025 매출",672,"","🟡 한경 2026.2.25"],
 ["FY2025 영업이익",207,"","OPM 30.8%"],
 ["D&A 가정(매출 4%)",{f:"B5*0.04"},"","🔴 실값 미확정(3~5% 레인지 중간)"],
 ["FY2025 EBITDA(추정)",{f:"B6+B7"},"","= 영업익 + D&A"],
 ["현재 시가총액",2700,"","🟡 FnGuide(시한성)"],
 ["",""],
 ["[Trading 대입 — EV/EBITDA] → 시사 EV(억원)","배수","시사 EV","함의 vs 현재시총"],
 ["EV/EBITDA 저 (8x, ConMed급)",8,{f:"B8*B12"},{f:'C12-B9&" 억"'}],
 ["EV/EBITDA 중 (13x, BSX/MMSI)",13,{f:"B8*B13"},{f:'C13-B9&" 억"'}],
 ["EV/EBITDA 고 (20x, Micro-Tech)",20,{f:"B8*B14"},{f:'C14-B9&" 억"'}],
 ["",""],
 ["[Trading 대입 — EV/EBIT] → 시사 EV(억원)","배수","시사 EV","함의 vs 현재시총"],
 ["  (EBIT = FY25 영업이익 207억)","","",""],
 ["EV/EBIT 저 (13x, ConMed)",13,{f:"B6*B18"},{f:'C18-B9&" 억"'}],
 ["EV/EBIT 중 (18.8x, BSX)",18.8,{f:"B6*B19"},{f:'C19-B9&" 억"'}],
 ["EV/EBIT 고 (22x, MMSI)",22,{f:"B6*B20"},{f:'C20-B9&" 억"'}],
 ["",""],
 ["[Transaction 대입 — EV/Revenue] → 시사 EV(억원)","배수","시사 EV","함의 vs 현재시총"],
 ["태웅 딜 (6.5x 매출)",6.5,{f:"B5*B23"},{f:'C23-B9&" 억"'}],
 ["Apollo/BSX시도 (8x 매출)",8,{f:"B5*B24"},{f:'C24-B9&" 억"'}],
 ["",""],
 ["▶ 결론: 현재 시총(~2,700억)은 트레이딩 하단~저트랜잭션 구간. 전략 인수자 프리미엄","","",""],
 ["  (태웅 6.5x·Apollo 8x 매출 → 4,300~5,400억)과 갭 = M&A/리레이팅 업사이드.","","",""],
 ["  ※ 멀티플·분모는 추정 — 면접에선 '구간'으로, basis(순현금·TTM/fwd) 명시.","","",""],
];
const ws3 = XLSX.utils.aoa_to_sheet(v);
ws3["!cols"]=[{wch:32},{wch:14},{wch:14},{wch:22}];
XLSX.utils.book_append_sheet(wb, ws3, "함의 밸류에이션");

// ===== Sheet 4: 출처 =====
const s = [
 ["출처 & 신뢰도","",""],
 ["🟢 HIGH","1차/공식 PR·공시","Olympus·BSX 공식 PR, 엠아이텍 DART 사업보고서"],
 ["🟡 MED","리포트·언론·aggregator","FnGuide 멀티플, stockanalysis/S&P/Yahoo, 트랜잭션 분모(타겟 재무)"],
 ["🔴 LOW","미확정","D&A 실값, 일부 단일소스 멀티플"],
 ["",""],
 ["트레이딩","BSX/MMSI/CNMD","stockanalysis.com (2026-06-24)"],
 ["","Micro-Tech 688029","finance.yahoo.com / gurufocus"],
 ["트랜잭션","Olympus→태웅","olympus-global.com PR / FierceBiotech $370M"],
 ["","BSX→엠아이텍(무산)","FierceBiotech / MedTech Dive"],
 ["","BSX→Apollo","news.bostonscientific.com PR ($615M, 매출$76M)"],
 ["","Merit→Cook/Biolife","MedTech Dive ($210M / $120M)"],
 ["섹터밴드","EBITDA/Revenue","FOCUS / MedDeviceGuide / Nelson Advisors 2025.9"],
 ["",""],
 ["주의: 멀티플·시총 시한성 → 면접 직전(6/25~26) 재확인. 딜 멀티플 분모는 보도추정.","",""],
];
const ws4 = XLSX.utils.aoa_to_sheet(s);
ws4["!cols"]=[{wch:14},{wch:22},{wch:52}];
XLSX.utils.book_append_sheet(wb, ws4, "출처");

XLSX.writeFile(wb, "MITech_Comps.xlsx");
console.log("comps workbook written: MITech_Comps.xlsx");
