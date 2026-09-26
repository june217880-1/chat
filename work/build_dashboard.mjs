import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outputDir = "/Users/hyeonsukbae/Documents/Codex/2026-09-26/https-docs-google-com-spreadsheets-d/outputs/01a0dbdb-acb1-7f62-8ba3-d68552e2b170";
const previewDir = "/Users/hyeonsukbae/Documents/Codex/2026-09-26/https-docs-google-com-spreadsheets-d/work/previews";
const outputPath = `${outputDir}/16개센터_실적관리_대시보드.xlsx`;

const centers = [
  {center:"서울",type:"단독",g25:"비교불가",r25:null,t25:null,jScore:51.749,aScore:48.480,jRank:12,aRank:21,jGrade:"A",aGrade:"A",target:"A",risk:"안전",rates:[0.878,0.073,0.585,0.463,0.146],scores:[17.811,11.686,6.272,6.893,5.818],focus:"알선취업률·연계실적",action:"직접알선 2건 이상 확보, 180일 유지 증빙 점검",basis:"2025 동일 운영단위 확인 불가"},
  {center:"성남",type:"컨소",g25:"A",r25:34,t25:57.516,jScore:45.681,aScore:45.679,jRank:93,aRank:75,jGrade:"B",aGrade:"B",target:"A",risk:"안전",rates:[0.615,0.168,0.162,0.234,0.269],scores:[14.901,13.499,4.810,5.046,7.423],focus:"A등급 진입·연계점수",action:"만족도와 연계점수 확정, 취업률 60% 이상 유지",basis:"스탭스 성남센터 직접 매칭"},
  {center:"청주",type:"단독",g25:"비교불가",r25:null,t25:null,jScore:46.597,aScore:45.617,jRank:63,aRank:79,jGrade:"B",aGrade:"B",target:"B",risk:"위험",rates:[0.944,0.000,0.417,0.278,0.111],scores:[18.657,10.510,5.649,5.445,5.356],focus:"알선취업률·소표본 변동",action:"신규 취업자를 직접알선으로 연결, 알선 2건 우선",basis:"2025 동일 운영단위 확인 불가"},
  {center:"구미",type:"컨소",g25:"B",r25:100,t25:55.138,jScore:45.010,aScore:44.839,jRank:117,aRank:109,jGrade:"B",aGrade:"B",target:"B",risk:"안전",rates:[0.586,0.146,0.176,0.246,0.277],scores:[14.382,13.103,4.885,4.931,7.537],focus:"취업률·연계점수",action:"취업률 60% 달성, 컨소 연계실적 조기 확정",basis:"스탭스 구미센터 직접 매칭"},
  {center:"인천북부",type:"컨소",g25:"B",r25:71,t25:55.803,jScore:44.919,aScore:44.703,jRank:121,aRank:119,jGrade:"B",aGrade:"B",target:"B",risk:"안전",rates:[0.616,0.104,0.176,0.189,0.258],scores:[15.282,12.350,4.889,4.900,7.282],focus:"현재 지표 유지·가점",action:"파트너별 실적 누락 점검, 만족도와 연계가점 확보",basis:"2025 동일 컨소 주관기관 기준"},
  {center:"광주",type:"컨소",g25:"D",r25:248,t25:50.584,jScore:43.466,aScore:43.568,jRank:191,aRank:182,jGrade:"C",aGrade:"C",target:"B",risk:"보통",rates:[0.578,0.121,0.144,0.170,0.264],scores:[14.284,12.648,4.733,4.546,7.357],focus:"취업률·고임금·연계",action:"중단 3명 사유 정리, 대체 실적과 취업률 60% 확보",basis:"스탭스 광주광산구센터 직접 매칭"},
  {center:"부천",type:"단독",g25:"B",r25:123,t25:54.590,jScore:43.187,aScore:43.223,jRank:207,aRank:201,jGrade:"C",aGrade:"C",target:"B",risk:"보통",rates:[0.617,0.046,0.145,0.209,0.253],scores:[14.949,11.242,4.679,5.137,7.216],focus:"알선취업률",action:"직접알선 취업자 5명 목표, 만족도 보완",basis:"스탭스 부천센터 직접 매칭"},
  {center:"서울동부",type:"컨소",g25:"C",r25:210,t25:52.724,jScore:43.111,aScore:43.047,jRank:212,aRank:209,jGrade:"C",aGrade:"C",target:"B",risk:"보통",rates:[0.583,0.046,0.157,0.248,0.276],scores:[14.379,11.308,4.787,5.062,7.512],focus:"알선취업률·가점",action:"가점 2점, 만족도, 파트너 역할별 알선실적 관리",basis:"스탭스 서울동부센터 직접 매칭"},
  {center:"서울남부",type:"컨소",g25:"C",r25:180,t25:53.367,jScore:42.692,aScore:42.839,jRank:228,aRank:217,jGrade:"C",aGrade:"C",target:"B",risk:"보통",rates:[0.583,0.047,0.185,0.213,0.256],scores:[14.568,11.327,4.923,4.764,7.258],focus:"알선취업률·고용유지",action:"직접알선 확대, 유지 증빙과 컨소 책임목표 확정",basis:"2025 동일 컨소 주관기관 기준"},
  {center:"구로",type:"단독",g25:"C",r25:212,t25:52.564,jScore:42.433,aScore:42.592,jRank:238,aRank:231,jGrade:"C",aGrade:"C",target:"B",risk:"위험",rates:[0.574,0.037,0.134,0.237,0.287],scores:[14.221,11.107,4.635,4.972,7.657],focus:"알선취업률·조기취업",action:"본사 알선지원, 단독기관 연계실적 동시 확보",basis:"스탭스 구로센터 직접 매칭"},
  {center:"대구동부",type:"컨소",g25:"C",r25:176,t25:53.450,jScore:41.911,aScore:43.059,jRank:256,aRank:208,jGrade:"D",aGrade:"C",target:"B",risk:"위험",rates:[0.572,0.100,0.127,0.195,0.253],scores:[14.246,12.284,4.677,4.635,7.218],focus:"취업률·고임금·연계",action:"취업률 60%와 연계점수 확보, 상승세 유지",basis:"2025 동일 컨소 주관기관 기준"},
  {center:"부산동부",type:"단독",g25:"C",r25:189,t25:53.227,jScore:41.753,aScore:42.797,jRank:258,aRank:221,jGrade:"D",aGrade:"C",target:"B",risk:"위험",rates:[0.582,0.131,0.138,0.175,0.219],scores:[14.140,12.621,4.650,4.617,6.769],focus:"취업률·유지·연계",action:"취업률 60%, 연계 2점 이상, 유지 증빙 점검",basis:"스탭스 부산동부센터 직접 매칭"},
  {center:"고양",type:"컨소",g25:"B",r25:84,t25:55.451,jScore:41.927,aScore:42.292,jRank:254,aRank:246,jGrade:"D",aGrade:"D",target:"B",risk:"위험",rates:[0.537,0.093,0.131,0.185,0.246],scores:[13.663,12.143,4.685,4.672,7.131],focus:"취업률·고임금",action:"먼저 C 경계 돌파, 취업자 집중 발굴과 연계 2점",basis:"스탭스 고양센터 직접 매칭"},
  {center:"전주",type:"컨소",g25:"C",r25:207,t25:52.856,jScore:42.129,aScore:41.931,jRank:248,aRank:251,jGrade:"D",aGrade:"D",target:"C",risk:"위험",rates:[0.553,0.054,0.143,0.180,0.264],scores:[13.779,11.453,4.726,4.609,7.365],focus:"취업률·알선·연계",action:"직접알선과 취업자를 동시 확대, 연계실적 확정",basis:"스탭스 전주센터 직접 매칭"},
  {center:"천안",type:"단독",g25:"B",r25:142,t25:54.324,jScore:40.954,aScore:40.661,jRank:274,aRank:277,jGrade:"D",aGrade:"D",target:"C",risk:"위험",rates:[0.501,0.074,0.088,0.189,0.211],scores:[13.115,11.707,4.462,4.712,6.665],focus:"취업률·조기취업·유지",action:"해외취업 인정 확인, 취업자 구조개선 집중",basis:"스탭스 천안센터 직접 매칭"},
  {center:"서영사무소",type:"단독",g25:"비교불가",r25:null,t25:null,jScore:37.932,aScore:38.113,jRank:287,aRank:287,jGrade:"D",aGrade:"D",target:"C",risk:"위험",rates:[0.428,0.034,0.074,0.126,0.188],scores:[12.026,11.054,4.404,4.268,6.360],focus:"취업률·알선·조기취업",action:"본사 개입형 취업·알선 파이프라인 주간 운영",basis:"2025 동일 운영단위 확인 불가"},
];

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Dashboard");
const detail = wb.worksheets.add("Center Detail");
const weekly = wb.worksheets.add("Weekly Plan");
const sources = wb.worksheets.add("Sources");

const C = {
  navy:"#17365D", blue:"#2F75B5", lightBlue:"#D9EAF7", teal:"#2A7F86",
  green:"#70AD47", lightGreen:"#E2F0D9", amber:"#F4B183", lightAmber:"#FFF2CC",
  red:"#C00000", lightRed:"#FCE4D6", gray:"#E7E6E6", dark:"#1F2937", white:"#FFFFFF",
};
const font = "Arial";
const headerFormat = {fill:C.navy,font:{name:font,size:10,bold:true,color:C.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
const sectionFormat = {fill:C.lightBlue,font:{name:font,size:10,bold:true,color:C.navy},borders:{preset:"outside",style:"thin",color:C.blue}};

for (const s of [dashboard, detail, weekly, sources]) {
  s.showGridLines = false;
}
dashboard.tabColor = C.navy;
weekly.tabColor = C.blue;
sources.tabColor = "#A5A5A5";

// Sources and threshold assumptions.
sources.getRange("A2").values = [["출처 및 기준"]];
sources.getRange("A2").format = {font:{name:font,size:14,bold:true,color:C.navy}};
sources.getRange("A4:B7").values = [
  ["항목","내용"],
  ["2026년 실적","https://docs.google.com/spreadsheets/d/1xmeOgfiZJBL_IRK86J0VZFSv63vJtGy8/edit?gid=1999365257"],
  ["2025년 최종실적","https://drive.google.com/file/d/1Xjzg_2W7kdvaQFuhmfxfz2ZTlRKhXLpd/view"],
  ["기준일","2026년 8월 말 실적, 파일 수정일 2026-09-26"],
];
sources.getRange("A4:B4").format = headerFormat;
sources.getRange("A5:B7").format.font = {name:font,size:10,color:C.dark};
sources.getRange("A10:B14").values = [
  ["8월 등급","점수 하한"], ["A",45.9929350079649], ["B",43.6690651278641], ["C",42.3585665458295], ["D",38.1133134316125],
];
sources.getRange("A10:B10").format = headerFormat;
sources.getRange("B11:B14").format.numberFormat = "0.000";
sources.getRange("D4:D7").values = [["분석 한계"],["6월 팩터별 세부점수는 제공 파일에 없어 6→8월 팩터 증감은 산출하지 않음"],["2025년 총점에는 만족도·가점이 포함되어 2026년 8월 총점과 직접 비교하지 않음"],["2025년 동일 운영단위 확인이 어려운 서울·청주·서영사무소는 비교불가로 표시"]];
sources.getRange("D4").format = headerFormat;
sources.getRange("D5:D7").format = {font:{name:font,size:10,color:C.dark},wrapText:true,verticalAlignment:"top"};
sources.getRange("A4:B14").format.borders = {preset:"outside",style:"thin",color:"#BFBFBF"};
sources.getRange("A:A").format.columnWidth = 22;
sources.getRange("B:B").format.columnWidth = 78;
sources.getRange("C:C").format.columnWidth = 3;
sources.getRange("D:D").format.columnWidth = 64;
sources.getRange("5:7").format.rowHeight = 38;

// Center detail: source inputs and calculated movement/gap fields.
const detailHeaders = ["센터","구분","2025 등급","2025 순위","2025 총점","6월 점수","8월 점수","점수 증감","6월 순위","8월 순위","순위 개선","6월 등급","8월 등급","목표 등급","목표등급 차이","원본 위험도","관리 구간","취업률","알선취업률","조기취업률","고임금률","고용유지율","취업 점수","알선 점수","조기 점수","고임금 점수","유지 점수","중점관리","즉시 실행과제","2025 매칭 기준"];
detail.getRange("A1:AD1").values = [detailHeaders];
detail.getRange("A1:AD1").format = headerFormat;
const detailRows = centers.map(c => [c.center,c.type,c.g25,c.r25,c.t25,c.jScore,c.aScore,null,c.jRank,c.aRank,null,c.jGrade,c.aGrade,c.target,null,c.risk,null,...c.rates,...c.scores,c.focus,c.action,c.basis]);
detail.getRange("A2:AD17").values = detailRows;
detail.getRange("H2").formulas = [["=G2-F2"]]; detail.getRange("H2:H17").fillDown();
detail.getRange("K2").formulas = [["=I2-J2"]]; detail.getRange("K2:K17").fillDown();
detail.getRange("O2").formulas = [["=MAX(0,IF(N2=\"A\",Sources!$B$11,IF(N2=\"B\",Sources!$B$12,IF(N2=\"C\",Sources!$B$13,Sources!$B$14)))-G2)"]]; detail.getRange("O2:O17").fillDown();
detail.getRange("Q2").formulas = [["=IF(O2=0,\"목표 충족\",IF(O2<=0.5,\"즉시 승급권\",IF(O2<=1.5,\"집중관리\",\"구조개선\")))"]]; detail.getRange("Q2:Q17").fillDown();
detail.getRange("D2:K17").format.numberFormat = "0.000";
detail.getRange("D2:D17").format.numberFormat = "0";
detail.getRange("I2:K17").format.numberFormat = "0";
detail.getRange("O2:O17").format.numberFormat = "0.000";
detail.getRange("R2:V17").format.numberFormat = "0.0%";
detail.getRange("W2:AA17").format.numberFormat = "0.000";
detail.getRange("A2:AD17").format = {font:{name:font,size:9,color:C.dark},verticalAlignment:"center"};
detail.getRange("AB2:AD17").format.wrapText = true;
detail.getRange("A1:AD17").format.borders = {insideHorizontal:{style:"thin",color:"#E7E6E6"},bottom:{style:"thin",color:"#BFBFBF"}};
detail.freezePanes.freezeRows(1); detail.freezePanes.freezeColumns(2);
detail.getRange("A:A").format.columnWidth=14; detail.getRange("B:B").format.columnWidth=9; detail.getRange("C:C").format.columnWidth=11;
detail.getRange("D:Q").format.columnWidth=11; detail.getRange("R:AA").format.columnWidth=12;
detail.getRange("AB:AB").format.columnWidth=24; detail.getRange("AC:AC").format.columnWidth=42; detail.getRange("AD:AD").format.columnWidth=32;
detail.getRange("2:17").format.rowHeight=34;
detail.getRange("H2:H17").conditionalFormats.add("colorScale",{colors:[C.lightRed,"#FFFFFF",C.lightGreen],thresholds:["min",{type:"percentile",value:50},"max"]});
detail.getRange("O2:O17").conditionalFormats.add("colorScale",{colors:[C.lightGreen,C.lightAmber,C.lightRed],thresholds:["min",{type:"percentile",value:50},"max"]});
detail.getRange("P2:P17").conditionalFormats.add("containsText",{text:"위험",format:{fill:C.lightRed,font:{bold:true,color:C.red}}});
detail.getRange("P2:P17").conditionalFormats.add("containsText",{text:"보통",format:{fill:C.lightAmber,font:{bold:true,color:"#9C6500"}}});
detail.getRange("M2:M17").conditionalFormats.add("containsText",{text:"D",format:{fill:C.lightRed,font:{bold:true,color:C.red}}});
detail.getRange("M2:M17").conditionalFormats.add("containsText",{text:"A",format:{fill:C.lightGreen,font:{bold:true,color:"#375623"}}});

// Weekly plan: 16 editable management rows.
const weeklyHeaders = ["센터","위험도","현재 등급","목표 등급","점수 차이","취업률","알선취업률","관리 구간","이번 주 핵심과제","증빙·가점 확인","담당자","마감일","진행상태","알선 목표(건)","주간 실적(건)","주간 메모"];
weekly.getRange("A1:P1").values = [weeklyHeaders]; weekly.getRange("A1:P1").format = headerFormat;
for(let r=2;r<=17;r++){
  weekly.getRange(`A${r}:I${r}`).formulas = [[`='Center Detail'!A${r}`,`='Center Detail'!P${r}`,`='Center Detail'!M${r}`,`='Center Detail'!N${r}`,`='Center Detail'!O${r}`,`='Center Detail'!R${r}`,`='Center Detail'!S${r}`,`='Center Detail'!Q${r}`,`='Center Detail'!AC${r}`]];
}
weekly.getRange("J2:J17").values = centers.map(c=>[c.type==="컨소"?"파트너별 연계·만족도·증빙":"연계·만족도·증빙"]);
weekly.getRange("K2:P17").values = centers.map(()=>[null,null,"미착수",null,null,null]);
weekly.getRange("M2:M17").dataValidation = {rule:{type:"list",values:["미착수","진행중","완료","지연"]}};
weekly.getRange("E2:E17").format.numberFormat="0.000"; weekly.getRange("F2:G17").format.numberFormat="0.0%";
weekly.getRange("L2:L17").format.numberFormat="yyyy-mm-dd";
weekly.getRange("A2:P17").format = {font:{name:font,size:9,color:C.dark},verticalAlignment:"center"};
weekly.getRange("I2:J17").format.wrapText=true;
weekly.getRange("K2:P17").format.fill=C.lightAmber;
weekly.getRange("A1:P17").format.borders={insideHorizontal:{style:"thin",color:"#E7E6E6"},bottom:{style:"thin",color:"#BFBFBF"}};
weekly.getRange("B2:B17").conditionalFormats.add("containsText",{text:"위험",format:{fill:C.lightRed,font:{bold:true,color:C.red}}});
weekly.getRange("M2:M17").conditionalFormats.add("containsText",{text:"완료",format:{fill:C.lightGreen,font:{bold:true,color:"#375623"}}});
weekly.getRange("M2:M17").conditionalFormats.add("containsText",{text:"지연",format:{fill:C.lightRed,font:{bold:true,color:C.red}}});
weekly.freezePanes.freezeRows(1); weekly.freezePanes.freezeColumns(1);
weekly.getRange("A:A").format.columnWidth=14; weekly.getRange("B:H").format.columnWidth=11; weekly.getRange("I:I").format.columnWidth=42; weekly.getRange("J:J").format.columnWidth=26;
weekly.getRange("K:K").format.columnWidth=13; weekly.getRange("L:L").format.columnWidth=13; weekly.getRange("M:O").format.columnWidth=12; weekly.getRange("P:P").format.columnWidth=30;
weekly.getRange("2:17").format.rowHeight=34;

// Dashboard title and KPI cards.
dashboard.getRange("A2").values = [["2026년 16개 센터 실적관리 대시보드"]];
dashboard.getRange("A2").format = {font:{name:font,size:16,bold:true,color:C.navy}};
dashboard.getRange("A3").values = [["8월 말 기준. 2027년 사업제안 전 등급 방어·상향과 주간 실행관리에 사용"]];
dashboard.getRange("A3").format = {font:{name:font,size:10,italic:true,color:"#666666"}};
dashboard.getRange("A4:O4").format.borders = {bottom:{style:"thin",color:C.blue}};

const cards = [
  {range:"A5:C7",label:"전체 센터",formula:"=COUNTA('Center Detail'!A2:A17)",fmt:"0"},
  {range:"E5:G7",label:"현재 A·B",formula:"=COUNTIF('Center Detail'!M2:M17,\"A\")+COUNTIF('Center Detail'!M2:M17,\"B\")",fmt:"0"},
  {range:"I5:K7",label:"목표등급 충족",formula:"=COUNTIF('Center Detail'!O2:O17,0)",fmt:"0"},
  {range:"M5:O7",label:"위험 센터",formula:"=COUNTIF('Center Detail'!P2:P17,\"위험\")",fmt:"0"},
  {range:"A9:C11",label:"평균 점수 증감",formula:"=AVERAGE('Center Detail'!H2:H17)",fmt:"0.000"},
  {range:"E9:G11",label:"평균 순위 개선",formula:"=AVERAGE('Center Detail'!K2:K17)",fmt:"0.0"},
  {range:"I9:K11",label:"취업률 60% 이상",formula:"=COUNTIF('Center Detail'!R2:R17,\">=0.6\")",fmt:"0"},
  {range:"M9:O11",label:"알선취업률 10% 이상",formula:"=COUNTIF('Center Detail'!S2:S17,\">=0.1\")",fmt:"0"},
];
for(const card of cards){
  const [start,end]=card.range.split(":");
  const startCol=start.match(/[A-Z]+/)[0], startRow=Number(start.match(/\d+/)[0]);
  const endCol=end.match(/[A-Z]+/)[0], endRow=Number(end.match(/\d+/)[0]);
  dashboard.mergeCells(`${startCol}${startRow}:${endCol}${startRow}`);
  dashboard.mergeCells(`${startCol}${startRow+1}:${endCol}${endRow}`);
  dashboard.getRange(`${startCol}${startRow}`).values=[[card.label]];
  dashboard.getRange(`${startCol}${startRow+1}`).formulas=[[card.formula]];
  dashboard.getRange(card.range).format={fill:"#F7F9FC",font:{name:font,size:10,color:C.dark},horizontalAlignment:"center",verticalAlignment:"center",borders:{preset:"outside",style:"thin",color:"#B4C7E7"}};
  dashboard.getRange(`${startCol}${startRow}`).format.font={name:font,size:10,bold:true,color:C.navy};
  dashboard.getRange(`${startCol}${startRow+1}`).format.font={name:font,size:16,bold:true,color:C.blue};
  dashboard.getRange(`${startCol}${startRow+1}`).format.numberFormat=card.fmt;
}

// Helper data and native charts.
dashboard.getRange("T5:U9").values = [["등급","센터 수"],["A",null],["B",null],["C",null],["D",null]];
for(let r=6;r<=9;r++) dashboard.getRange(`U${r}`).formulas=[[`=COUNTIF('Center Detail'!M2:M17,T${r})`]];
dashboard.getRange("T12:V28").values = [["센터","6월 점수","8월 점수"],...centers.map(c=>[c.center,c.jScore,c.aScore])];
dashboard.getRange("U13:V28").format.numberFormat="0.0";
const gradeChart = dashboard.charts.add("doughnut",dashboard.getRange("T5:U9"));
gradeChart.title="8월 등급 분포"; gradeChart.titleTextStyle.fontSize=12; gradeChart.titleTextStyle.typeface=font;
gradeChart.legend={position:"right",textStyle:{typeface:font,fontSize:10}}; gradeChart.setPosition("A14","H28");
const scoreChart = dashboard.charts.add("bar",dashboard.getRange("T12:V28"));
scoreChart.title="센터별 6월·8월 합계점수"; scoreChart.titleTextStyle.fontSize=12; scoreChart.titleTextStyle.typeface=font;
scoreChart.legend={position:"top",textStyle:{typeface:font,fontSize:10}};
scoreChart.xAxis={axisType:"textAxis",textStyle:{typeface:font,fontSize:9}};
scoreChart.yAxis={numberFormatCode:"0.0",numberFormatSourceLinked:false,textStyle:{typeface:font,fontSize:9}};
scoreChart.setPosition("J14","R31");

dashboard.getRange("A31:H31").values=[["우선관리 센터","8월 등급","목표","점수 차이","취업률","알선취업률","관리 구간","핵심 조치"]];
dashboard.getRange("A31:H31").format=headerFormat;
const priorityNames=["서영사무소","천안","고양","구로","부산동부","서울남부","서울동부","전주"];
for(let i=0;i<priorityNames.length;i++){
  const row=i+32, drow=centers.findIndex(c=>c.center===priorityNames[i])+2;
  dashboard.getRange(`A${row}:H${row}`).formulas=[[
    `='Center Detail'!A${drow}`,`='Center Detail'!M${drow}`,`='Center Detail'!N${drow}`,`='Center Detail'!O${drow}`,
    `='Center Detail'!R${drow}`,`='Center Detail'!S${drow}`,`='Center Detail'!Q${drow}`,`='Center Detail'!AC${drow}`
  ]];
}
dashboard.getRange("D32:D39").format.numberFormat="0.000"; dashboard.getRange("E32:F39").format.numberFormat="0.0%";
dashboard.getRange("A32:H39").format={font:{name:font,size:9,color:C.dark},verticalAlignment:"center"};
dashboard.getRange("H32:H39").format.wrapText=true; dashboard.getRange("32:39").format.rowHeight=32;
dashboard.getRange("A31:H39").format.borders={insideHorizontal:{style:"thin",color:"#E7E6E6"},bottom:{style:"thin",color:"#BFBFBF"}};

dashboard.mergeCells("J34:R34");
dashboard.getRange("J34").values=[["경영진 판단"]]; dashboard.getRange("J34:R34").format=sectionFormat;
const insights = [
  "1. 8월 등급은 D 6개에서 4개로 개선됐지만 A·B는 5개로 동일합니다.",
  "2. 16개 중 취업률 60% 이상 5개, 알선취업률 10% 이상 6개입니다.",
  "3. 광주·성남·부천·대구동부·서울동부는 목표등급 경계에 가깝습니다.",
  "4. 천안·고양은 2025년 대비 등급 하락폭이 커 본사 개입이 필요합니다.",
  "5. 서영사무소는 가점만으로 목표달성이 어려워 취업·알선 구조개선이 필요합니다.",
];
for(let i=0;i<insights.length;i++){
  const row=35+i;
  dashboard.mergeCells(`J${row}:R${row}`);
  dashboard.getRange(`J${row}`).values=[[insights[i]]];
}
dashboard.getRange("J35:R39").format={font:{name:font,size:10,color:C.dark},verticalAlignment:"center",wrapText:true,borders:{preset:"outside",style:"thin",color:"#D9E2F3"}};
dashboard.getRange("35:39").format.rowHeight=28;

dashboard.getRange("A:A").format.columnWidth=15; dashboard.getRange("B:C").format.columnWidth=9; dashboard.getRange("D:D").format.columnWidth=4;
dashboard.getRange("E:G").format.columnWidth=11; dashboard.getRange("H:H").format.columnWidth=38; dashboard.getRange("I:I").format.columnWidth=4;
dashboard.getRange("J:R").format.columnWidth=12; dashboard.getRange("T:V").format.columnWidth=12;
dashboard.getRange("T:V").format.font={name:font,size:8,color:"#808080"};
dashboard.getRange("T:V").format.fill="#F2F2F2";

// General font and alignment finalization.
for (const s of [dashboard, detail, weekly, sources]) {
  const used = s.getUsedRange();
  if (used) used.format.verticalAlignment = "center";
}

wb.recalculate();
await fs.mkdir(outputDir,{recursive:true});
await fs.mkdir(previewDir,{recursive:true});

const check1=await wb.inspect({kind:"table",range:"Dashboard!A1:R39",include:"values,formulas",tableMaxRows:39,tableMaxCols:18,maxChars:12000});
console.log("DASHBOARD_CHECK\n"+check1.ndjson);
const check2=await wb.inspect({kind:"table",range:"Center Detail!A1:AD17",include:"values,formulas",tableMaxRows:17,tableMaxCols:30,maxChars:12000});
console.log("DETAIL_CHECK\n"+check2.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:100},summary:"formula error scan"});
console.log("ERROR_SCAN\n"+errors.ndjson);

for(const [sheetName,range,fileName] of [["Dashboard","A1:R39","dashboard.png"],["Center Detail","A1:AD17","center-detail.png"],["Weekly Plan","A1:P17","weekly-plan.png"],["Sources","A1:D14","sources.png"]]){
  const blob=await wb.render({sheetName,range,scale:1,format:"png"});
  await fs.writeFile(`${previewDir}/${fileName}`,new Uint8Array(await blob.arrayBuffer()));
}

const out=await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(`OUTPUT ${outputPath}`);
