import { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { createRoot } from "react-dom/client";

function smsTemplate(caseId, product, status, delivery) {
  return "[무역알림] 고객님 안녕하세요.\n주문번호 " + caseId + " (" + product + ") 상태가 [" + status + "](으)로 업데이트되었습니다.\n예상 납기: " + (delivery || "미정") + "\n문의: 관리자 연락처";
}


// ============================================================
// 数据结构 v3 — 按真实调查书/报价单流程重构
// QuoteCase 报价案件
//   └── items[] 产品（一个案件可含多个产品）
//         └── sources[] 比价行（一个产品可有多个供应商链接/报价）
// 每个比价行包含：链接、单价、内陆费、包装、一箱数量、箱数、CBM、重量、
//                 纳期、微信ID、打样费、LOGO费等（对应报价单全部列）
// ============================================================

const LIVE_RATES = { CNY: 190.5, USD: 1382.0, KRW: 1 };
const CURRENCIES = ["CNY", "USD", "KRW"];
const CURRENCY_SYMBOL = { CNY: "¥", USD: "$", KRW: "₩" };

const PERMISSIONS = {
  admin: {
    label: { zh: "管理员", ko: "관리자" },
    viewPurchase: true, editPurchase: true,
    viewSale: true, editSale: true,
    viewProfit: true,
    editRate: true, editRateMode: true,
    canDelete: true, ownOnly: false,
  },
  sales: {
    label: { zh: "营业组", ko: "영업팀" },
    viewPurchase: true, editPurchase: true,
    viewSale: true, editSale: true,
    viewProfit: false,
    editRate: true, editRateMode: false,
    canDelete: true, ownOnly: false,
  },
  purchasing: {
    label: { zh: "采购组", ko: "구매팀" },
    viewPurchase: true, editPurchase: true,
    viewSale: false, editSale: false,
    viewProfit: false,
    editRate: true, editRateMode: false,
    canDelete: true, ownOnly: false,
  },
  // 兼职：不能删除任何案件；只能编辑自己担任负责人的案件，其余只读
  parttime: {
    label: { zh: "兼职", ko: "파트타임" },
    viewPurchase: true, editPurchase: true,
    viewSale: true, editSale: true,
    viewProfit: false,
    editRate: true, editRateMode: false,
    canDelete: false, ownOnly: true,
  },
};

// 兼职是否可编辑此案件：必须是该案件的营业负责人或报价负责人
function canTouchCase(perm, user, c) {
  if (!perm.ownOnly) return true;
  return [c.salesRep, c.purchaseRep].includes(user.name) || [c.salesRep, c.purchaseRep].includes(user.id);
}

// ============================================================
// 员工账号 — 管理员在「员工管理」页添加员工并分组
// 员工用姓名或ID登录，登录后权限自动按所属组生效
// ⚠ 原型仅作演示：正式系统必须由后端验证（密码加密+会话/JWT），
//   前端判断可被绕过，不能作为真正的安全边界
// ============================================================
// ⚠ 원형 데모: 비밀번호는 평문으로 저장됩니다.
// 정식 시스템에서는 반드시 bcrypt 해시 + HTTPS + 서버측 세션/JWT 인증 필요
const MOCK_USERS = [
  // 관리층
  { id: "admin",   name: "김기남", role: "admin",      pw: "admin123" },
  { id: "eunji",   name: "김은진", role: "admin",      pw: "eunji2026" },
  // 영업팀
  { id: "eunjae",  name: "이은재", role: "sales",      pw: "sales01" },
  { id: "heekyung",name: "김희경", role: "sales",      pw: "sales02" },
  // 구매팀
  { id: "wujin",   name: "왕우진", role: "purchasing", pw: "purchase01" },
  // 파트타임
  { id: "tmp01",   name: "최아라", role: "parttime",   pw: "temp2026" },
];

const T = {
  zh: {
    appTitle: "中韩贸易订单管理系统",
    dashboard: "首页", quoteSystem: "报价系统", orderSystem: "下单系统",
    totalQuotes: "报价案件", totalOrders: "下单案件",
    urgentCases: "即将到期", overdueCases: "已逾期",
    recentUpdated: "最近更新案件",
    newQuote: "新增报价案件", newOrder: "新增下单",
    convertToOrder: "转为下单", detail: "详情/编辑",
    allStatus: "全部状态", allCustomers: "全部客户", allStaff: "全部负责人",
    edit: "编辑", delete: "删除", save: "保存", cancel: "取消", close: "关闭",
    aiTranslate: "AI翻译", copyResult: "复制", translating: "翻译中...",
    caseNo: "案件编号", orderNo: "订单编号", linkedQuote: "关联报价",
    customer: "客户名", contact: "联系方式", supplier: "供应商/工厂",
    product: "产品名", spec: "规格", qty: "数量",
    items: "产品", sources: "比价", addItem: "+ 添加产品", addSource: "+ 添加比价行",
    detailQty: "详细数量", colorSize: "颜色/尺寸", sampleNeeded: "样品",
    customerReq: "客户要求(特殊事项)", supplierReply: "供应商回复",
    url: "网址链接", unitPrice: "单价", totalPrice: "总金额",
    inlandFee: "内陆费", packing: "包装方式", perBox: "一箱数量", boxCount: "箱数",
    cbm: "体积CBM", weightKg: "重量kg", origin: "原产地", factory: "工厂",
    cbmCalc: "CBM计算器", cbmL: "长(cm)", cbmW: "宽(cm)", cbmH: "高(cm)",
    cbmQty: "每箱数量", cbmUnit: "单个", cbmBox: "整箱", cbmResult: "计算结果", cbmApply: "应用到表格",
    leadTime: "纳期", sellerContact: "卖家联系方式", sampleNote: "打样说明",
    attachments: "附件", uploadFiles: "上传文件", attachHint: "图片 / 厂家报价单 / 资料 / 视频", noFiles: "暂无附件",
    moldFee: "版费", sampleTime: "打样时间", sourceNote: "备注",
    selected: "选定", selectThis: "选此报价", noQuote: "未报价/不接单",
    purchaseCurrency: "币种", purchaseAmount: "采购价(原币)",
    exchangeRate: "汇率", lockedRate: "本单汇率(手动输入)", rateHint: "参考",
    purchaseKRW: "采购价(韩币)", salePrice: "销售价 KRW",
    profit: "利润 KRW", profitRate: "利润率",
    landedCost: "到手成本(含内陆费)",
    rateAuto: "实时汇率", rateManual: "固定汇率(管理员)",
    rateSettings: "汇率设置", rateMode: "汇率模式",
    noPermission: "权限不足",
    quoteDate: "报价日期", quoteValidity: "报价有效期",
    estimatedDelivery: "预计交期", factoryDelivery: "工厂交期",
    note: "备注", responsible: "负责人",
    salesRep: "营业负责人", purchaseRep: "报价负责人",
    category: "分类", catDaily: "日常生活用品", catMachine: "机器类", allCategories: "全部",
    searchPlaceholder: "🔍 搜索关键词：客户 / 产品 / 工厂 / 备注...",
    readOnlyOwn: "只读：仅本人负责的案件可编辑", status: "状态",
    smsBtn: "📱 发送短信通知", smsSend: "发送", smsCancel: "取消",
    smsTitle: "短信通知预览", smsPhone: "客户手机号",
    smsSent: "✓ 短信已发送", smsPending: "※ 演示模式：实际发送需接入短信API（SENS/솔라피）",
    smsTemplate: null,
    days: "天", overdue: "逾期",
    paymentStatus: "付款状态", logistics: "物流备注", customerNote: "客户备注",
    arrivalDate: "韩国到港日", customsClear: "通关状态", deliveryDate: "货物送达日",
    shipInfo: "船期/航班", warehouseDate: "中国仓库入库日",
    loadingDate: "装货日", estShipDate: "预计出库日",
    exportExcel: "导出 Excel", exportNote: "导出内容已按当前角色权限过滤",
    exportOrderList: "📊 发注清单导出", exportQuoteSheet: "📋 报价单导出",
    exportChoose: "选择导出格式", exportAll: "全部案件", exportFiltered: "当前筛选",
    sheetOrderList: "发注清单", sheetQuoteSummary: "报价汇总", sheetPriceCompare: "比价明细",
    exporting: "导出中...",
    currentRole: "当前角色",
    loginTitle: "员工登录", enterNameOrId: "员工ID / 姓名", enterPw: "密码",
    loginBtn: "登录", logout: "退出登录",
    userNotFound: "未找到该账号，请联系管理员", wrongPw: "密码错误",
    showPw: "显示", hidePw: "隐藏",
    staffMgmt: "员工管理", addStaff: "添加员工", userId: "员工ID", userName: "姓名", userRole: "所属组", userPw: "初始密码",
    resetPw: "重置密码", resetPwDone: "密码已重置为: ",
    demoAccounts: "演示账号", staffExists: "该ID已存在", deleteSelf: "不能删除自己",
    staffHint: "管理员在此添加员工并分组，员工即可用姓名或ID登录",
    basicInfo: "基本信息", moneyInfo: "金额信息", logisticsInfo: "物流日程",
    compareInfo: "供应商比价", bestPrice: "最低价",
    selectFirst: "请先在比价表中选定一个报价",
    yes: "要", no: "不要",
    qs1: "调查书", qs2: "询价中", qs3: "已报价", qsInq: "客户问询", qsUpd: "更新报价", qs4: "客户确认",
    packingFee: "包装费",
    quoteOverdue: "报价超时", dayUnit: "天",
    inquiryLog: "问询记录", addInquiry: "+ 添加问询", inquiryContent: "问询/新要求内容",
    markInquiry: "客户问询", markUpdated: "完成更新报价", markConfirm: "客户确认",
    os1: "下单申请", os2: "已下单", os3: "工厂出库", os4: "中国仓库",
    os5: "装货", os6: "韩国到港", os7: "通关中", os8: "清关完成", os9: "货物送达",
    lang: "한국어",
  },
  ko: {
    appTitle: "한중 무역 주문 관리 시스템",
    dashboard: "홈", quoteSystem: "견적 시스템", orderSystem: "주문 시스템",
    totalQuotes: "견적 건수", totalOrders: "주문 건수",
    urgentCases: "마감 임박", overdueCases: "기한 초과",
    recentUpdated: "최근 업데이트 건",
    newQuote: "견적 케이스 추가", newOrder: "주문 추가",
    convertToOrder: "주문 전환", detail: "상세/편집",
    allStatus: "전체 상태", allCustomers: "전체 고객", allStaff: "전체 담당자",
    edit: "편집", delete: "삭제", save: "저장", cancel: "취소", close: "닫기",
    aiTranslate: "AI 번역", copyResult: "복사", translating: "번역 중...",
    caseNo: "케이스 번호", orderNo: "주문 번호", linkedQuote: "연결 견적",
    customer: "고객명", contact: "연락처", supplier: "공급업체/공장",
    product: "제품명", spec: "규격", qty: "수량",
    items: "제품", sources: "비교 견적", addItem: "+ 제품 추가", addSource: "+ 비교 견적 추가",
    detailQty: "상세수량", colorSize: "색상/사이즈", sampleNeeded: "샘플",
    customerReq: "고객 요구사항(특이사항)", supplierReply: "공급업체 회신",
    url: "사이트주소", unitPrice: "단가", totalPrice: "총금액",
    inlandFee: "내륙비", packing: "포장", perBox: "박스당", boxCount: "박스수",
    cbm: "CBM", weightKg: "중량kg", origin: "원산지", factory: "공장",
    cbmCalc: "CBM 계산기", cbmL: "가로(cm)", cbmW: "세로(cm)", cbmH: "높이(cm)",
    cbmQty: "박스당 수량", cbmUnit: "낱개", cbmBox: "박스 전체", cbmResult: "계산 결과", cbmApply: "표에 적용",
    leadTime: "납기일", sellerContact: "판매자 연락처", sampleNote: "견본 설명",
    attachments: "첨부파일", uploadFiles: "파일 업로드", attachHint: "이미지 / 공장 견적서 / 자료 / 영상", noFiles: "첨부 없음",
    moldFee: "판비(版费)", sampleTime: "타양시간", sourceNote: "비고",
    selected: "선정", selectThis: "이 견적 선택", noQuote: "미견적/거절",
    purchaseCurrency: "통화", purchaseAmount: "매입가(원통화)",
    exchangeRate: "환율", lockedRate: "건별 환율(직접 입력)", rateHint: "참고",
    purchaseKRW: "매입가(KRW)", salePrice: "판매가 KRW",
    profit: "이익 KRW", profitRate: "이익률",
    landedCost: "도착 원가(내륙비 포함)",
    rateAuto: "실시간 환율", rateManual: "고정 환율(관리자)",
    rateSettings: "환율 설정", rateMode: "환율 모드",
    noPermission: "권한 부족",
    quoteDate: "견적일", quoteValidity: "견적 유효기간",
    estimatedDelivery: "예상 납기", factoryDelivery: "공장 납기",
    note: "비고", responsible: "담당자",
    salesRep: "영업 담당자", purchaseRep: "견적 담당자",
    category: "분류", catDaily: "생활용품", catMachine: "기계류", allCategories: "전체",
    searchPlaceholder: "🔍 키워드 검색: 고객 / 제품 / 공장 / 비고...",
    readOnlyOwn: "읽기 전용: 본인 담당 건만 편집 가능", status: "상태",
    smsBtn: "📱 SMS 알림 발송", smsSend: "발송", smsCancel: "취소",
    smsTitle: "SMS 미리보기", smsPhone: "고객 전화번호",
    smsSent: "✓ 발송 완료", smsPending: "※ 데모 모드: 실제 발송은 SMS API 연동 필요 (SENS / 솔라피)",
    smsTemplate: null,
    days: "일", overdue: "기한 초과",
    paymentStatus: "결제 상태", logistics: "물류 비고", customerNote: "고객 비고",
    arrivalDate: "한국 입항일", customsClear: "통관 상태", deliveryDate: "배송 완료일",
    shipInfo: "선편/항공편", warehouseDate: "중국 창고 입고일",
    loadingDate: "선적일", estShipDate: "예상 출고일",
    exportExcel: "엑셀 내보내기", exportNote: "현재 역할 권한에 따라 필터링됨",
    exportOrderList: "📊 발주 리스트 내보내기", exportQuoteSheet: "📋 견적서 내보내기",
    exportChoose: "내보내기 형식 선택", exportAll: "전체 건", exportFiltered: "현재 필터",
    sheetOrderList: "발주 리스트", sheetQuoteSummary: "견적 요약", sheetPriceCompare: "비교 견적 상세",
    exporting: "내보내는 중...",
    currentRole: "현재 역할",
    loginTitle: "직원 로그인", enterNameOrId: "직원 ID / 이름", enterPw: "비밀번호",
    loginBtn: "로그인", logout: "로그아웃",
    userNotFound: "계정을 찾을 수 없습니다. 관리자에게 문의하세요", wrongPw: "비밀번호가 틀렸습니다",
    showPw: "표시", hidePw: "숨기기",
    staffMgmt: "직원 관리", addStaff: "직원 추가", userId: "직원 ID", userName: "이름", userRole: "소속 그룹", userPw: "초기 비밀번호",
    resetPw: "비밀번호 초기화", resetPwDone: "비밀번호가 초기화되었습니다: ",
    demoAccounts: "데모 계정", staffExists: "이미 존재하는 ID입니다", deleteSelf: "본인은 삭제할 수 없습니다",
    staffHint: "관리자가 직원을 추가하고 그룹을 지정하면, 직원은 이름 또는 ID로 로그인합니다",
    basicInfo: "기본 정보", moneyInfo: "금액 정보", logisticsInfo: "물류 일정",
    compareInfo: "공급업체 비교 견적", bestPrice: "최저가",
    selectFirst: "비교 표에서 견적을 먼저 선택하세요",
    yes: "필요", no: "불필요",
    qs1: "조사서", qs2: "견적 요청 중", qs3: "견적 완료", qsInq: "고객 문의", qsUpd: "견적 갱신", qs4: "고객 확인",
    packingFee: "포장비",
    quoteOverdue: "견적 지연", dayUnit: "일",
    inquiryLog: "문의 기록", addInquiry: "+ 문의 추가", inquiryContent: "문의/추가 요구사항",
    markInquiry: "고객 문의", markUpdated: "견적 갱신 완료", markConfirm: "고객 확인",
    os1: "주문 신청", os2: "주문 완료", os3: "공장 출고", os4: "중국 창고",
    os5: "선적", os6: "한국 입항", os7: "통관 중", os8: "통관 완료", os9: "배송 완료",
    lang: "中文",
  }
};

// 报价状态流：调查书 → 询价中 → 已报价 → (客户问询 → 更新报价 →)* → 客户确认
// 问询/更新报价 是可循环的中间状态：客户追问时回到问询，重新询价后进入更新报价
const QUOTE_STATUSES = ["qs1","qs2","qs3","qsInq","qsUpd","qs4"];
const ORDER_STATUSES = ["os1","os2","os3","os4","os5","os6","os7","os8","os9"];

// ============================================================
// Mock Data — 基于真实调查书案例「三角饭团机」
// ============================================================
let _sid = 100;
const sid = () => `s${_sid++}`;

const MOCK_QUOTES = [
  {
    id: "E-260602", customer: "이종민", contact: "010-XXXX-XXXX",
    category: "machine",
    salesRep: "김민준", purchaseRep: "王小芳",
    status: "qsInq", quoteDate: "2026-06-02", statusDate: "2026-06-09",
    quoteValidity: "2026-06-20", estimatedDelivery: "2026-07-10",
    inquiryLog: [
      { date: "2026-06-09", content: "견적가에 배송비 포함인지 확인 요망. 그리고 380V 60Hz 버전도 견적 받아주세요." },
    ],
    items: [
      {
        id: "i1", productName: "삼각밥성형기",
        detailQty: "삼각밥성형기 1대", colorSize: "", sampleNeeded: false,
        customerReq: "하루 2~3시간 작업 기준 200개 정도. 220V60Hz 사용 가능 여부, 전기용량, 사이즈/무게, 납기일 확인 요망",
        salePrice: 1450000,
        sources: [
          { id: sid(), url: "https://www.alibaba.com/product-detail/Hot-Sale-Discount-Onigiri...", factory: "대형 생산라인 회사", unitPrice: "", qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: "", packing: "", packingFee: "", perBox: "", boxCount: "", cbm: "", weightKg: "", origin: "", leadTime: "", wechatId: "", sampleNote: "", attachments: [], note: "양이 너무 적어서 안함 (거절)", selected: false, noQuote: true },
          { id: sid(), url: "https://www.alibaba.com/product-detail/Used-FUJISEIKI-Triangle...", factory: "일본 FUJISEIKI 중고", unitPrice: 850000, qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: 0, packing: "", packingFee: "", perBox: "", boxCount: "", cbm: "", weightKg: "", origin: "일본", leadTime: "2달", wechatId: "", sampleNote: "", attachments: [], note: "중고기계, 현재 재고 없음. 신품 발주시 약 80만위안, 시간당 1800개", selected: false, noQuote: false },
          { id: sid(), url: "(2번 링크 동일 판매자)", factory: "반자동 소형", unitPrice: 7200, qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: 0, packing: "스펀지(폼)+종이박스 44×29.5×31cm", packingFee: 0, perBox: 1, boxCount: 1, cbm: 0.04, weightKg: 7, origin: "중국", leadTime: "7근무일", wechatId: "", sampleNote: "", attachments: [], note: "김밥 사이즈 76×76×30mm, 220V60Hz, 식품급 POM 재료", selected: true, noQuote: false },
          { id: sid(), url: "https://www.alibaba.com/product-detail/Onigiri-Rice-Ball-Processing...", factory: "Luohe High Mechanical", unitPrice: 3770, qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: 0, packing: "목상자 500×355×455mm", packingFee: 100, perBox: 1, boxCount: 1, cbm: 0.04, weightKg: 21, origin: "중국", leadTime: "5~7일", wechatId: "", sampleNote: "샘플비 ¥300, 대량 발주 시 전액 감면 가능", attachments: [], note: "220V60Hz, 스테인리스 304", selected: false, noQuote: false },
        ]
      },
      {
        id: "i2", productName: "낱개김포장기계",
        detailQty: "낱개김포장기계 1대", colorSize: "", sampleNeeded: false,
        customerReq: "100매짜리 낱개 삼각김밥김 포장 가능해야 함",
        salePrice: 1900000,
        sources: [
          { id: sid(), url: "(상동 2번 판매자)", factory: "반자동 소형", unitPrice: 8000, qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: 0, packing: "스펀지(폼)+종이박스 45.5×33.5×32cm", packingFee: 0, perBox: 1, boxCount: 1, cbm: 0.04, weightKg: 12, origin: "중국", leadTime: "7근무일", wechatId: "", sampleNote: "", attachments: [], note: "220V60Hz", selected: true, noQuote: false },
          { id: sid(), url: "(상동 Luohe)", factory: "Luohe High Mechanical", unitPrice: 6580, qty: 1, currency: "CNY", lockedRate: 190.5, inlandFee: 0, packing: "목상자 500×340×485mm", packingFee: 100, perBox: 1, boxCount: 1, cbm: 0.08, weightKg: 31, origin: "중국", leadTime: "5~7일", wechatId: "", sampleNote: "", attachments: [], note: "220V60Hz, 스테인리스 304", selected: false, noQuote: false },
        ]
      }
    ]
  },
  {
    // 演示「报价超时」黄色提醒：询价中状态停留超过2天
    id: "E-260605", customer: "박상우", contact: "010-XXXX-XXXX",
    category: "daily",
    salesRep: "이수진", purchaseRep: "박철수",
    status: "qs2", quoteDate: "2026-06-05", statusDate: "2026-06-05",
    quoteValidity: "", estimatedDelivery: "2026-07-25",
    inquiryLog: [],
    items: [
      {
        id: "i1", productName: "스테인리스 주방 선반",
        detailQty: "1200×450×1800mm, 4단", colorSize: "실버", sampleNeeded: false,
        customerReq: "SUS201 가능, 조립식 선호",
        salePrice: "",
        sources: [
          { id: sid(), url: "https://detail.1688.com/offer/...", factory: "佛山不锈钢厂", unitPrice: "", qty: 10, currency: "CNY", lockedRate: 190.5, inlandFee: "", packing: "", packingFee: "", perBox: "", boxCount: "", cbm: "", weightKg: "", origin: "", leadTime: "", wechatId: "", sampleNote: "", attachments: [], note: "회신 대기 중", selected: false, noQuote: false },
        ]
      }
    ]
  },
];

const MOCK_ORDERS = [
  { id: "O2026-001", linkedQuote: "E-260601", customer: "현대무역", product: "스테인리스 볼트 M8×30", spec: "SUS304, 1000개/박스", qty: 50, purchaseCurrency: "CNY", purchaseAmount: 2250, lockedRate: 189.2, inlandFee: 120, salePrice: 12000000, supplier: "上海钢铁件有限公司", salesRep: "이은재", purchaseRep: "왕우진", customerPhone: "010-1234-5678", status: "os4", factoryDelivery: "2026-06-18", warehouseDate: "2026-06-22", loadingDate: "", shipInfo: "", arrivalDate: "", deliveryDate: "2026-07-05", paymentStatus: "입금완료", logistics: "중국 창고 입고 확인됨", customerNote: "포장 상태 확인 요망", estShipDate: "2026-06-20", customsClear: "" },
  { id: "O2026-002", linkedQuote: "", customer: "삼성물산", product: "LED 조명 패널 300×300mm", spec: "24W, 6500K", qty: 200, purchaseCurrency: "USD", purchaseAmount: 2160, lockedRate: 1379.5, inlandFee: 0, salePrice: 4400000, supplier: "广州光电科技", salesRep: "김희경", purchaseRep: "왕우진", customerPhone: "010-9876-5432", status: "os6", factoryDelivery: "2026-06-10", warehouseDate: "2026-06-14", loadingDate: "2026-06-18", shipInfo: "COSCO SHIPPING", arrivalDate: "2026-06-28", deliveryDate: "2026-07-05", paymentStatus: "미결제", logistics: "인천항 입항 완료", customerNote: "", estShipDate: "2026-06-12", customsClear: "" },
];

// ============================================================
// 计算
// 到手成本 = (单价×数量 + 内陆费 + 打样费 + LOGO费) × 汇率
// ============================================================
// 到手成本 = (单价×数量 + 内陆费 + 包装费) × 汇率
// 打样费一般大货可减免，改为「打样说明」文字栏，由采购组填写说明
function calcSourceKRW(s) {
  const base = (parseFloat(s.unitPrice) || 0) * (parseFloat(s.qty) || 1)
    + (parseFloat(s.inlandFee) || 0) + (parseFloat(s.packingFee) || 0);
  return Math.round(base * (parseFloat(s.lockedRate) || 1));
}
function calcItemMoney(item) {
  const sel = item.sources.find(s => s.selected);
  const purchaseKRW = sel ? calcSourceKRW(sel) : 0;
  const sale = parseFloat(item.salePrice) || 0;
  const profit = sale - purchaseKRW;
  const profitRate = sale > 0 ? (profit / sale) * 100 : 0;
  return { sel, purchaseKRW, profit, profitRate };
}
function calcOrderMoney(o) {
  const purchaseKRW = Math.round(((parseFloat(o.purchaseAmount) || 0) + (parseFloat(o.inlandFee) || 0)) * (parseFloat(o.lockedRate) || 1));
  const sale = parseFloat(o.salePrice) || 0;
  const profit = sale - purchaseKRW;
  const profitRate = sale > 0 ? (profit / sale) * 100 : 0;
  return { purchaseKRW, profit, profitRate };
}
function getDaysLeft(d) { if (!d) return null; return Math.ceil((new Date(d) - new Date()) / 86400000); }
function urgencyOf(days) { if (days === null) return "normal"; if (days < 0) return "overdue"; if (days <= 5) return "urgent"; return "normal"; }

// ============================================================
// 通用小组件
// ============================================================
function StatusBadge({ status, lang }) {
  const t = T[lang];
  const COLOR_MAP = {
    qs1: ["#e8f0fe", "#1a56db"], qs2: ["#e8f0fe", "#1a56db"], qs3: ["#ede9fe", "#6d28d9"],
    qsInq: ["#fef3c7", "#b45309"], qsUpd: ["#cffafe", "#0e7490"], qs4: ["#e6f4ea", "#1a7f37"],
  };
  let bg, color;
  if (COLOR_MAP[status]) { [bg, color] = COLOR_MAP[status]; }
  else {
    const list = ORDER_STATUSES;
    const idx = list.indexOf(status);
    bg = "#e8f0fe"; color = "#1a56db";
    if (status === "os9") { bg = "#e6f4ea"; color = "#1a7f37"; }
    else if (idx >= list.length * 0.66) { bg = "#fce8f3"; color = "#9d174d"; }
    else if (idx >= list.length * 0.33) { bg = "#fef9c3"; color = "#854d0e"; }
  }
  return <span style={{ background: bg, color, borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 500, whiteSpace: "nowrap" }}>{t[status]}</span>;
}

function DaysBadge({ days, lang }) {
  const t = T[lang];
  if (days === null) return null;
  let bg = "#f0fdf4", color = "#166534", label = `D-${days}`;
  if (days < 0) { bg = "#fee2e2"; color = "#b91c1c"; label = `${t.overdue} ${Math.abs(days)}${t.days}`; }
  else if (days <= 5) { bg = "#fef9c3"; color = "#92400e"; }
  return <span style={{ background: bg, color, borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 600 }}>{label}</span>;
}

function ProgressBar({ status, lang, showLabels }) {
  const isOrder = status.startsWith("os");
  const list = isOrder ? ORDER_STATUSES : QUOTE_STATUSES;
  const idx = list.indexOf(status);
  const t = lang ? T[lang] : null;
  return (
    <div style={{ display: "flex", gap: 4, marginTop: 8 }}>
      {list.map((s, i) => (
        <div key={s} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center" }}>
          <div style={{ width: "100%", height: 3, background: i <= idx ? "#1a56db" : "#e5e7eb", borderRadius: 2 }} />
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: i === idx ? "#1a56db" : i < idx ? "#60a5fa" : "#d1d5db", marginTop: 2, boxShadow: i === idx ? "0 0 0 3px rgba(26,86,219,0.18)" : "none" }} />
          {showLabels && t && (
            <span style={{ fontSize: 9, marginTop: 3, color: i === idx ? "#1a56db" : i < idx ? "#64748b" : "#cbd5e1", fontWeight: i === idx ? 700 : 400, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "100%" }}>
              {t[s]}
            </span>
          )}
        </div>
      ))}
    </div>
  );
}

function Lock({ lang }) {
  return <span style={{ color: "#d1d5db", fontSize: 12, fontStyle: "italic" }}>🔒 {T[lang].noPermission}</span>;
}

// ============================================================
// AI翻译
// ============================================================
function AITranslateBtn({ text, lang, fieldLabel }) {
  const [open, setOpen] = useState(false);
  const [result, setResult] = useState("");
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const t = T[lang];
  const targetLang = lang === "zh" ? "ko" : "zh";

  async function handleTranslate() {
    if (!text) return;
    setOpen(true); setLoading(true); setResult("");
    try {
      const systemPrompt = targetLang === "ko"
        ? "你是专业的中韩贸易翻译。将中文翻译为自然的韩语，保留专业术语。只输出翻译结果。"
        : "당신은 한중 무역 전문 번역가입니다. 한국어를 자연스러운 중국어로 번역하세요. 번역 결과만 출력하세요.";
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, system: systemPrompt, messages: [{ role: "user", content: text }] })
      });
      const data = await res.json();
      setResult(data.content?.map(b => b.text || "").join("") || "번역 실패");
    } catch { setResult("翻译失败 / 번역 실패"); }
    setLoading(false);
  }

  return (
    <span>
      <button onClick={handleTranslate} style={{ fontSize: 11, padding: "1px 8px", borderRadius: 5, border: "1px solid #3b82f6", background: "#eff6ff", color: "#1d4ed8", cursor: "pointer", marginLeft: 4 }}>
        🌐 {lang === "zh" ? "AI韩译" : "AI한→중"}
      </button>
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 2000, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setOpen(false)}>
          <div style={{ background: "#fff", borderRadius: 12, padding: 24, width: 400, boxShadow: "0 8px 32px rgba(0,0,0,0.18)" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 12 }}>{t.aiTranslate} — {fieldLabel}</div>
            <div style={{ background: "#f9fafb", borderRadius: 8, padding: 10, fontSize: 13, marginBottom: 12, maxHeight: 80, overflow: "auto" }}>{text}</div>
            <div style={{ background: "#eff6ff", borderRadius: 8, padding: 10, fontSize: 13, minHeight: 50, maxHeight: 120, overflow: "auto" }}>
              {loading ? <span style={{ color: "#9ca3af" }}>{t.translating}</span> : result}
            </div>
            <div style={{ display: "flex", gap: 8, marginTop: 14, justifyContent: "flex-end" }}>
              {result && !loading && (
                <button onClick={() => { navigator.clipboard.writeText(result); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
                  style={{ fontSize: 12, padding: "5px 14px", borderRadius: 7, border: "1px solid #10b981", background: copied ? "#dcfce7" : "#fff", color: "#059669", cursor: "pointer" }}>
                  {copied ? "✓" : t.copyResult}
                </button>
              )}
              <button onClick={() => setOpen(false)} style={{ fontSize: 12, padding: "5px 14px", borderRadius: 7, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{t.close}</button>
            </div>
          </div>
        </div>
      )}
    </span>
  );
}

// ============================================================
// 报价案件卡片
// ============================================================
// ============================================================
// 报价超时判断：状态停在 调查书/询价中/客户问询 超过2天 → 黄色提醒
// ============================================================
function getQuoteOverdueDays(q) {
  if (!["qs1", "qs2", "qsInq"].includes(q.status)) return 0;
  const ref = q.statusDate || q.quoteDate;
  if (!ref) return 0;
  const days = Math.floor((new Date() - new Date(ref)) / 86400000);
  return days > 2 ? days : 0;
}

function QuoteCard({ q, lang, perm, user, onDetail, onConvert, onDelete, onQuickStatus }) {
  const t = T[lang];
  const editable = canTouchCase(perm, user, q);
  const days = getDaysLeft(q.estimatedDelivery);
  const overdueDays = getQuoteOverdueDays(q);
  const u = overdueDays > 0 ? "urgent" : urgencyOf(days);
  const borderColor = u === "overdue" ? "#fecaca" : u === "urgent" ? "#fde68a" : "#e5e7eb";
  const totalSources = q.items.reduce((s, it) => s + it.sources.length, 0);
  const allSelected = q.items.every(it => it.sources.some(s => s.selected));
  return (
    <div style={{ background: "#fff", border: `1.5px solid ${borderColor}`, borderRadius: 14, padding: "16px 20px", marginBottom: 14, boxShadow: u !== "normal" ? "0 2px 8px rgba(234,179,8,0.12)" : "0 1px 4px rgba(0,0,0,0.06)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ minWidth: 0 }}>
          <span style={{ fontSize: 11, color: "#9ca3af", fontFamily: "monospace" }}>{q.id}</span>
          {q.category && (
            <span style={{ fontSize: 10, marginLeft: 8, padding: "1px 8px", borderRadius: 5, background: q.category === "machine" ? "#ede9fe" : "#ecfdf5", color: q.category === "machine" ? "#6d28d9" : "#047857", fontWeight: 600 }}>
              {q.category === "machine" ? `⚙️ ${t.catMachine}` : `🧺 ${t.catDaily}`}
            </span>
          )}
          <div style={{ fontWeight: 600, fontSize: 15, marginTop: 2 }}>{q.customer}</div>
          <div style={{ fontSize: 13, color: "#374151", marginTop: 2 }}>
            {q.items.map(it => it.productName).join(" · ")}
          </div>
          <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 4 }}>
            📦 {q.items.length} {t.items} &nbsp;|&nbsp; 🔍 {totalSources} {t.sources}
            {q.inquiryLog?.length > 0 && <>&nbsp;|&nbsp; 💬 {q.inquiryLog.length} {t.inquiryLog}</>}
          </div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
          <StatusBadge status={q.status} lang={lang} />
          {overdueDays > 0 && (
            <span style={{ background: "#fef3c7", color: "#b45309", borderRadius: 6, padding: "2px 8px", fontSize: 11, fontWeight: 700 }}>
              ⚠ {t.quoteOverdue} {overdueDays}{t.dayUnit}
            </span>
          )}
          <DaysBadge days={days} lang={lang} />
        </div>
      </div>

      {/* 每个产品的选定报价摘要 */}
      <div style={{ marginTop: 10, padding: "8px 12px", background: "#f8fafc", borderRadius: 8 }}>
        {q.items.map(it => {
          const { sel, purchaseKRW, profit, profitRate } = calcItemMoney(it);
          return (
            <div key={it.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 12, padding: "3px 0", gap: 8 }}>
              <span style={{ color: "#374151", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{it.productName}</span>
              <span style={{ whiteSpace: "nowrap", color: "#6b7280" }}>
                {sel
                  ? <>✓ {CURRENCY_SYMBOL[sel.currency]}{(parseFloat(sel.unitPrice)||0).toLocaleString()} → <b style={{ color: "#0369a1" }}>₩{purchaseKRW.toLocaleString()}</b>
                      {perm.viewProfit && <span style={{ color: profit >= 0 ? "#166534" : "#dc2626", marginLeft: 6 }}>({profitRate.toFixed(0)}%)</span>}
                    </>
                  : <span style={{ color: "#d97706" }}>⚠ 미선정</span>}
              </span>
            </div>
          );
        })}
      </div>

      <div style={{ display: "flex", gap: 14, marginTop: 8, fontSize: 12, color: "#6b7280", flexWrap: "wrap" }}>
        <span title={t.salesRep}>👤 {t.salesRep}: {q.salesRep || "—"}</span>
        <span title={t.purchaseRep}>📝 {t.purchaseRep}: {q.purchaseRep || "—"}</span>
        <span>📅 {q.quoteDate}</span>
      </div>
      {/* 状态进度条 + 状态名称标签 */}
      <ProgressBar status={q.status} lang={lang} showLabels />
      <div style={{ display: "flex", gap: 8, marginTop: 14, justifyContent: "flex-end", flexWrap: "wrap", alignItems: "center" }}>
        {!editable && <span style={{ fontSize: 11, color: "#d97706", marginRight: "auto" }}>🔒 {t.readOnlyOwn}</span>}
        {/* 状态流转快捷按钮：已报价/客户确认 → 客户问询 → 更新报价 → 客户确认 */}
        {editable && ["qs3", "qs4", "qsUpd"].includes(q.status) && (
          <button onClick={() => onQuickStatus(q.id, "qsInq")}
            style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #fcd34d", background: "#fffbeb", color: "#b45309", cursor: "pointer" }}>
            💬 {t.markInquiry}
          </button>
        )}
        {editable && q.status === "qsInq" && (
          <button onClick={() => onQuickStatus(q.id, "qsUpd")}
            style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #67e8f9", background: "#ecfeff", color: "#0e7490", cursor: "pointer" }}>
            ✦ {t.markUpdated}
          </button>
        )}
        {editable && ["qs3", "qsUpd"].includes(q.status) && (
          <button onClick={() => onQuickStatus(q.id, "qs4")}
            style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #86efac", background: "#f0fdf4", color: "#166534", cursor: "pointer" }}>
            ✓ {t.markConfirm}
          </button>
        )}
        {editable && q.status === "qs4" && (
          <button onClick={() => onConvert(q)} disabled={!allSelected} title={!allSelected ? t.selectFirst : ""}
            style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "none", background: allSelected ? "#1a56db" : "#cbd5e1", color: "#fff", cursor: allSelected ? "pointer" : "not-allowed" }}>
            {t.convertToOrder}
          </button>
        )}
        <button onClick={() => onDetail(q)} style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #93c5fd", background: "#eff6ff", color: "#1d4ed8", cursor: "pointer" }}>{t.detail}</button>
        {perm.canDelete && (
          <button onClick={() => onDelete(q.id)} style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #fecaca", background: "#fff5f5", color: "#dc2626", cursor: "pointer" }}>{t.delete}</button>
        )}
      </div>
    </div>
  );
}

// ============================================================
// 比价表（一个产品下的多行供应商报价）— 对应报价单全部列
// ============================================================
const SOURCE_COLS = [
  ["url", "url", 200], ["factory", "factory", 110],
  ["unitPrice", "unitPrice", 80], ["qty", "qty", 50],
  ["inlandFee", "inlandFee", 70], ["packing", "packing", 130], ["packingFee", "packingFee", 65],
  ["perBox", "perBox", 55], ["boxCount", "boxCount", 50],
  ["weightKg", "weightKg", 60],
  ["origin", "origin", 70], ["leadTime", "leadTime", 70],
  ["sampleNote", "sampleNote", 160], ["note", "sourceNote", 180],
];
const NUM_KEYS = ["unitPrice","qty","inlandFee","packingFee","perBox","boxCount","cbm","weightKg"];

// ============================================================
// 附件管理：图片/厂家报价单/资料/视频，支持预览
// 原型阶段文件保存在浏览器内存(ObjectURL)，刷新后消失；正式系统需上传到服务器/云存储
// ============================================================
function AttachBtn({ attachments = [], onChange, lang, label }) {
  const [open, setOpen] = useState(false);
  const t = T[lang];
  function handleFiles(e) {
    const files = Array.from(e.target.files || []);
    const added = files.map(f => ({
      id: `a${Date.now()}_${Math.random().toString(36).slice(2,7)}`,
      name: f.name, type: f.type, size: f.size, url: URL.createObjectURL(f)
    }));
    onChange([...attachments, ...added]);
    e.target.value = "";
  }
  function removeAtt(id) { onChange(attachments.filter(a => a.id !== id)); }
  const fmtSize = n => n > 1048576 ? (n/1048576).toFixed(1) + "MB" : Math.round(n/1024) + "KB";

  return (
    <>
      <button onClick={() => setOpen(true)}
        style={{ fontSize: 11, padding: "3px 9px", borderRadius: 6, border: "1px solid " + (attachments.length ? "#86efac" : "#d1d5db"), background: attachments.length ? "#f0fdf4" : "#f9fafb", color: attachments.length ? "#166534" : "#6b7280", cursor: "pointer", whiteSpace: "nowrap" }}>
        📎 {attachments.length || ""}
      </button>
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1500, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setOpen(false)}>
          <div style={{ background: "#fff", borderRadius: 14, padding: 24, width: 600, maxHeight: "82vh", overflowY: "auto", boxShadow: "0 12px 40px rgba(0,0,0,0.18)" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 4 }}>📎 {t.attachments}{label ? ` — ${label}` : ""}</div>
            <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 14 }}>{t.attachHint}</div>

            <label style={{ display: "block", border: "2px dashed #93c5fd", borderRadius: 10, padding: "18px 0", textAlign: "center", cursor: "pointer", background: "#f8fafc", marginBottom: 16, color: "#1d4ed8", fontSize: 13 }}>
              ⬆ {t.uploadFiles}
              <input type="file" multiple accept="image/*,video/*,.pdf,.xlsx,.xls,.doc,.docx,.zip" onChange={handleFiles} style={{ display: "none" }} />
            </label>

            {attachments.length === 0 && <div style={{ textAlign: "center", color: "#cbd5e1", fontSize: 12, padding: 16 }}>{t.noFiles}</div>}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 12 }}>
              {attachments.map(a => (
                <div key={a.id} style={{ border: "1px solid #e5e7eb", borderRadius: 10, overflow: "hidden", position: "relative", background: "#fafafa" }}>
                  <button onClick={() => removeAtt(a.id)}
                    style={{ position: "absolute", top: 4, right: 4, zIndex: 2, width: 20, height: 20, borderRadius: "50%", border: "none", background: "rgba(0,0,0,0.55)", color: "#fff", fontSize: 11, cursor: "pointer", lineHeight: "20px", padding: 0 }}>✕</button>
                  {a.type.startsWith("image/") ? (
                    <a href={a.url} target="_blank" rel="noreferrer">
                      <img src={a.url} alt={a.name} style={{ width: "100%", height: 100, objectFit: "cover", display: "block" }} />
                    </a>
                  ) : a.type.startsWith("video/") ? (
                    <video src={a.url} controls style={{ width: "100%", height: 100, objectFit: "cover", display: "block", background: "#000" }} />
                  ) : (
                    <a href={a.url} target="_blank" rel="noreferrer" style={{ display: "flex", alignItems: "center", justifyContent: "center", height: 100, fontSize: 30, textDecoration: "none" }}>
                      {a.name.match(/\.(xlsx?|csv)$/i) ? "📊" : a.name.match(/\.pdf$/i) ? "📕" : a.name.match(/\.(docx?)$/i) ? "📄" : "📁"}
                    </a>
                  )}
                  <div style={{ padding: "5px 8px", fontSize: 10, color: "#475569", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={a.name}>
                    {a.name}<br/><span style={{ color: "#9ca3af" }}>{fmtSize(a.size)}</span>
                  </div>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
              <button onClick={() => setOpen(false)} style={{ padding: "7px 18px", borderRadius: 8, border: "none", background: "#1a56db", color: "#fff", cursor: "pointer", fontSize: 13 }}>{t.close}</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}


// ============================================================
// CBM 계산기 — 장×폭×높이×수량으로 CBM 자동계산
// ============================================================
function CbmCalcBtn({ currentCbm, onApply, lang, canEdit }) {
  const t = T[lang];
  const [open, setOpen] = useState(false);
  const [dims, setDims] = useState({ l: "", w: "", h: "", qty: 1, mode: "box" });
  const d = (k, v) => setDims(prev => ({ ...prev, [k]: v }));

  // 단위: cm → m 변환, CBM = L×W×H / 1,000,000
  const singleCBM = (parseFloat(dims.l)||0) * (parseFloat(dims.w)||0) * (parseFloat(dims.h)||0) / 1000000;
  const totalCBM  = dims.mode === "box"
    ? singleCBM                              // 박스 1개 치수
    : singleCBM * (parseFloat(dims.qty)||1); // 낱개 × 수량

  function apply() {
    onApply(parseFloat(totalCBM.toFixed(4)));
    setOpen(false);
  }

  return (
    <>
      <button onClick={() => canEdit && setOpen(true)} title={t.cbmCalc}
        style={{ fontSize: 10, padding: "2px 6px", borderRadius: 4, border: "1px solid #d1d5db",
          background: "#f9fafb", color: "#6b7280", cursor: canEdit ? "pointer" : "default",
          whiteSpace: "nowrap" }}>
        📐
      </button>
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 2000,
          display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setOpen(false)}>
          <div style={{ background: "#fff", borderRadius: 14, padding: 24, width: 340,
            boxShadow: "0 12px 40px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 16 }}>📐 {t.cbmCalc}</div>

            {/* 모드 선택 */}
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              {[["box", t.cbmBox], ["unit", t.cbmUnit]].map(([m, label]) => (
                <button key={m} onClick={() => d("mode", m)}
                  style={{ flex: 1, padding: "7px 0", borderRadius: 8, fontSize: 12, cursor: "pointer",
                    border: dims.mode === m ? "2px solid #1a56db" : "1px solid #d1d5db",
                    background: dims.mode === m ? "#eff6ff" : "#f9fafb",
                    color: dims.mode === m ? "#1a56db" : "#6b7280",
                    fontWeight: dims.mode === m ? 700 : 400 }}>
                  {label}
                </button>
              ))}
            </div>

            {/* 치수 입력 */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, marginBottom: 12 }}>
              {[["l", t.cbmL], ["w", t.cbmW], ["h", t.cbmH]].map(([k, label]) => (
                <div key={k}>
                  <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 4 }}>{label}</label>
                  <input type="number" min="0" step="0.1" value={dims[k]}
                    onChange={e => d(k, e.target.value)}
                    style={{ width: "100%", padding: "7px 8px", borderRadius: 7,
                      border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
                </div>
              ))}
            </div>

            {/* 낱개 모드: 수량 입력 */}
            {dims.mode === "unit" && (
              <div style={{ marginBottom: 12 }}>
                <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.cbmQty}</label>
                <input type="number" min="1" value={dims.qty} onChange={e => d("qty", e.target.value)}
                  style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
              </div>
            )}

            {/* 계산 결과 */}
            <div style={{ background: "#f0f9ff", borderRadius: 9, padding: "12px 14px", marginBottom: 14 }}>
              <div style={{ fontSize: 11, color: "#0369a1", marginBottom: 6 }}>{t.cbmResult}</div>
              {dims.mode === "unit" && (
                <div style={{ fontSize: 11, color: "#64748b", marginBottom: 4 }}>
                  {t.cbmUnit}: {singleCBM.toFixed(6)} m³ × {dims.qty || 1} =
                </div>
              )}
              <div style={{ fontSize: 22, fontWeight: 700, color: "#0369a1" }}>
                {totalCBM.toFixed(4)} <span style={{ fontSize: 13, fontWeight: 400 }}>m³ (CBM)</span>
              </div>
              {currentCbm !== undefined && currentCbm !== "" && (
                <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 4 }}>
                  {lang === "zh" ? "当前值" : "현재값"}: {currentCbm}
                </div>
              )}
            </div>

            <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
              <button onClick={() => setOpen(false)}
                style={{ padding: "8px 18px", borderRadius: 8, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer", fontSize: 13 }}>
                {t.cancel}
              </button>
              <button onClick={apply} disabled={!totalCBM}
                style={{ padding: "8px 18px", borderRadius: 8, border: "none",
                  background: totalCBM ? "#1a56db" : "#cbd5e1", color: "#fff",
                  fontWeight: 600, fontSize: 13, cursor: totalCBM ? "pointer" : "not-allowed" }}>
                {t.cbmApply}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function SourceTable({ item, lang, perm, onChange }) {
  const t = T[lang];
  const canEdit = perm.editPurchase;
  const validSources = item.sources.filter(s => !s.noQuote && parseFloat(s.unitPrice) > 0);
  const minKRW = validSources.length ? Math.min(...validSources.map(calcSourceKRW)) : null;

  function setSource(sIdx, k, v) {
    const sources = item.sources.map((s, i) => i === sIdx ? { ...s, [k]: v } : s);
    onChange({ ...item, sources });
  }
  function selectSource(sIdx) {
    const sources = item.sources.map((s, i) => ({ ...s, selected: i === sIdx }));
    onChange({ ...item, sources });
  }
  function addSource() {
    onChange({ ...item, sources: [...item.sources, { id: sid(), url: "", factory: "", unitPrice: "", qty: 1, currency: "CNY", lockedRate: LIVE_RATES.CNY, inlandFee: "", packing: "", packingFee: "", perBox: "", boxCount: "", cbm: "", weightKg: "", origin: "", leadTime: "", wechatId: "", contactFiles: [], sampleNote: "", attachments: [], note: "", selected: false, noQuote: false }] });
  }
  function delSource(sIdx) {
    onChange({ ...item, sources: item.sources.filter((_, i) => i !== sIdx) });
  }

  return (
    <div style={{ overflowX: "auto", border: "1px solid #e5e7eb", borderRadius: 10, marginTop: 8 }}>
      <table style={{ borderCollapse: "collapse", fontSize: 12, minWidth: 1600 }}>
        <thead>
          <tr style={{ background: "#f1f5f9" }}>
            <th style={{ padding: "8px 8px", position: "sticky", left: 0, background: "#f1f5f9", zIndex: 1, minWidth: 60 }}>{t.selected}</th>
            {SOURCE_COLS.map(([k, lk, w]) => (
              <th key={k} style={{ padding: "8px 8px", textAlign: "left", whiteSpace: "nowrap", minWidth: w, color: "#475569", fontWeight: 600 }}>{t[lk]}</th>
            ))}
            <th style={{ padding: "8px 8px", minWidth: 80, color: "#475569" }}>{t.cbm} 📐</th>
            <th style={{ padding: "8px 8px", minWidth: 130, color: "#475569" }}>{t.sellerContact}</th>
            <th style={{ padding: "8px 8px", minWidth: 90, color: "#475569" }}>{t.exchangeRate}</th>
            <th style={{ padding: "8px 8px", minWidth: 100, color: "#0369a1", fontWeight: 700 }}>{t.landedCost}</th>
            <th style={{ padding: "8px 8px", minWidth: 55, color: "#475569" }}>{t.attachments}</th>
            <th style={{ minWidth: 40 }}></th>
          </tr>
        </thead>
        <tbody>
          {item.sources.map((s, sIdx) => {
            const krw = calcSourceKRW(s);
            const isBest = minKRW !== null && krw === minKRW && !s.noQuote && parseFloat(s.unitPrice) > 0;
            return (
              <tr key={s.id} style={{ borderTop: "1px solid #f1f5f9", background: s.selected ? "#eff6ff" : s.noQuote ? "#fafafa" : "#fff" }}>
                <td style={{ padding: "6px 8px", textAlign: "center", position: "sticky", left: 0, background: s.selected ? "#eff6ff" : "#fff", zIndex: 1 }}>
                  {s.noQuote
                    ? <span style={{ fontSize: 10, color: "#9ca3af" }}>{t.noQuote}</span>
                    : <input type="radio" checked={!!s.selected} onChange={() => canEdit && selectSource(sIdx)} disabled={!canEdit} style={{ cursor: canEdit ? "pointer" : "default" }} />}
                </td>
                {SOURCE_COLS.map(([k]) => (
                  <td key={k} style={{ padding: "4px 6px" }}>
                    <input type={NUM_KEYS.includes(k) ? "number" : "text"} value={s[k] ?? ""} disabled={!canEdit}
                      onChange={e => setSource(sIdx, k, e.target.value)}
                      title={s[k]}
                      style={{ width: "100%", padding: "4px 6px", borderRadius: 5, border: "1px solid #e5e7eb", fontSize: 12, boxSizing: "border-box", background: canEdit ? "#fff" : "#f9fafb", color: s.noQuote ? "#9ca3af" : "#1f2937" }} />
                  </td>
                ))}
                <td style={{ padding: "4px 6px" }}>
                  <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                    <input type="number" step="0.0001" value={s.cbm ?? ""} disabled={!canEdit}
                      onChange={e => setSource(sIdx, "cbm", e.target.value)}
                      style={{ width: 54, padding: "4px 5px", borderRadius: 5, border: "1px solid #e5e7eb", fontSize: 12, background: canEdit ? "#fff" : "#f9fafb" }} />
                    <CbmCalcBtn currentCbm={s.cbm} lang={lang} canEdit={canEdit}
                      onApply={val => setSource(sIdx, "cbm", val)} />
                  </div>
                </td>
                <td style={{ padding: "4px 6px" }}>
                  <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
                    <input type="text" value={s.wechatId ?? ""} disabled={!canEdit} placeholder="WeChat/电话/QQ..."
                      onChange={e => setSource(sIdx, "wechatId", e.target.value)}
                      style={{ flex: 1, minWidth: 0, padding: "4px 6px", borderRadius: 5, border: "1px solid #e5e7eb", fontSize: 12, background: canEdit ? "#fff" : "#f9fafb" }} />
                    <AttachBtn attachments={s.contactFiles || []} onChange={atts => setSource(sIdx, "contactFiles", atts)} lang={lang} label={t.sellerContact} />
                  </div>
                </td>
                <td style={{ padding: "4px 6px" }}>
                  <div style={{ display: "flex", gap: 3, alignItems: "center" }}>
                    <select value={s.currency} disabled={!canEdit} onChange={e => { setSource(sIdx, "currency", e.target.value); }}
                      style={{ padding: "4px 2px", borderRadius: 5, border: "1px solid #e5e7eb", fontSize: 11 }}>
                      {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                    <input type="number" step="0.01" value={s.lockedRate ?? ""} disabled={!canEdit}
                      onChange={e => setSource(sIdx, "lockedRate", e.target.value)}
                      style={{ width: 58, padding: "4px 5px", borderRadius: 5, border: "1.5px solid #93c5fd", fontSize: 12 }} />
                  </div>
                </td>
                <td style={{ padding: "4px 8px", fontWeight: 700, color: isBest ? "#166534" : "#0369a1", whiteSpace: "nowrap" }}>
                  {s.noQuote ? "—" : <>₩{krw.toLocaleString()}{isBest && <span style={{ fontSize: 10, marginLeft: 4, background: "#dcfce7", color: "#166534", padding: "1px 5px", borderRadius: 4 }}>{t.bestPrice}</span>}</>}
                </td>
                <td style={{ padding: "4px 6px", textAlign: "center" }}>
                  <AttachBtn attachments={s.attachments || []} onChange={atts => setSource(sIdx, "attachments", atts)} lang={lang} label={s.factory || s.url?.slice(0, 24)} />
                </td>
                <td style={{ padding: "4px 6px", textAlign: "center" }}>
                  {canEdit && <button onClick={() => delSource(sIdx)} style={{ border: "none", background: "none", color: "#dc2626", cursor: "pointer", fontSize: 13 }}>✕</button>}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      {canEdit && (
        <button onClick={addSource} style={{ margin: 8, fontSize: 12, padding: "5px 14px", borderRadius: 7, border: "1px dashed #93c5fd", background: "#f8fafc", color: "#1d4ed8", cursor: "pointer" }}>
          {t.addSource}
        </button>
      )}
    </div>
  );
}

// ============================================================
// 报价案件详情弹窗（调查书 + 比价 + 销售价）
// ============================================================
function QuoteDetailModal({ initial, lang, perm, user, onSave, onClose }) {
  const t = T[lang];
  const locked = initial ? !canTouchCase(perm, user, initial) : false;
  const [form, setForm] = useState(initial || {
    id: `E-${new Date().toISOString().slice(2,10).replace(/-/g,"")}`,
    customer: "", contact: "", category: "daily",
    salesRep: user?.name || "", purchaseRep: "", status: "qs1",
    quoteDate: new Date().toISOString().slice(0,10), quoteValidity: "", estimatedDelivery: "",
    statusDate: new Date().toISOString().slice(0,10), inquiryLog: [],
    items: [{ id: `i${Date.now()}`, productName: "", detailQty: "", colorSize: "", sampleNeeded: false, customerReq: "", salePrice: "", attachments: [], sources: [] }]
  });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v, ...(k === "status" ? { statusDate: new Date().toISOString().slice(0,10) } : {}) }));
  const setItem = (idx, item) => setForm(f => ({ ...f, items: f.items.map((it, i) => i === idx ? item : it) }));
  const addItem = () => setForm(f => ({ ...f, items: [...f.items, { id: `i${Date.now()}`, productName: "", detailQty: "", colorSize: "", sampleNeeded: false, customerReq: "", salePrice: "", attachments: [], sources: [] }] }));
  const delItem = idx => setForm(f => ({ ...f, items: f.items.filter((_, i) => i !== idx) }));
  const [newInquiry, setNewInquiry] = useState("");
  function addInquiry() {
    if (!newInquiry.trim()) return;
    setForm(f => ({
      ...f,
      inquiryLog: [...(f.inquiryLog || []), { date: new Date().toISOString().slice(0,10), content: newInquiry.trim() }],
      status: "qsInq", statusDate: new Date().toISOString().slice(0,10)
    }));
    setNewInquiry("");
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 16, padding: 26, width: "92vw", maxWidth: 1280, maxHeight: "92vh", overflowY: "auto", boxShadow: "0 16px 48px rgba(0,0,0,0.18)" }} onClick={e => e.stopPropagation()}>
        <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 16 }}>{initial ? t.detail : t.newQuote} — {form.id}</div>

        {locked && (
          <div style={{ background: "#fffbeb", border: "1px solid #fcd34d", borderRadius: 9, padding: "8px 14px", fontSize: 12, color: "#b45309", marginBottom: 14 }}>
            🔒 {t.readOnlyOwn}
          </div>
        )}
        <fieldset disabled={locked} style={{ border: "none", padding: 0, margin: 0 }}>
        {/* 案件头：客户信息（对应调查书顶部） */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: "10px 14px", marginBottom: 16 }}>
          <div>
            <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{t.category}</label>
            <select value={form.category || "daily"} onChange={e => set("category", e.target.value)}
              style={{ width: "100%", padding: "6px 6px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12 }}>
              <option value="daily">🧺 {t.catDaily}</option>
              <option value="machine">⚙️ {t.catMachine}</option>
            </select>
          </div>
          {[["customer", t.customer], ["contact", t.contact], ["salesRep", t.salesRep], ["purchaseRep", t.purchaseRep], ["quoteDate", t.quoteDate], ["quoteValidity", t.quoteValidity], ["estimatedDelivery", t.estimatedDelivery]].map(([k, label]) => (
            <div key={k}>
              <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{label}</label>
              <input type={["quoteDate","quoteValidity","estimatedDelivery"].includes(k) ? "date" : "text"} value={form[k] || ""} onChange={e => set(k, e.target.value)}
                style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12, boxSizing: "border-box" }} />
            </div>
          ))}
        </div>
        <div style={{ display: "flex", gap: 14, alignItems: "center", marginBottom: 14 }}>
          <label style={{ fontSize: 12, color: "#6b7280" }}>{t.status}</label>
          <select value={form.status} onChange={e => set("status", e.target.value)} style={{ padding: "6px 12px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12 }}>
            {QUOTE_STATUSES.map(s => <option key={s} value={s}>{t[s]}</option>)}
          </select>
        </div>

        {/* 问询记录：客户报价后的追问/新要求都记在这里，不用新建调查书 */}
        <div style={{ border: "1.5px solid #fde68a", borderRadius: 12, padding: "12px 16px", marginBottom: 18, background: "#fffdf5" }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: "#92400e", marginBottom: 8 }}>💬 {t.inquiryLog} ({(form.inquiryLog || []).length})</div>
          {(form.inquiryLog || []).map((iq, i) => (
            <div key={i} style={{ display: "flex", gap: 10, padding: "6px 0", borderBottom: "1px dashed #fde68a", fontSize: 12, alignItems: "flex-start" }}>
              <span style={{ color: "#b45309", fontFamily: "monospace", whiteSpace: "nowrap" }}>{iq.date}</span>
              <span style={{ flex: 1, color: "#374151" }}>{iq.content}</span>
              <AITranslateBtn text={iq.content} lang={lang} fieldLabel={t.inquiryContent} />
            </div>
          ))}
          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            <input value={newInquiry} onChange={e => setNewInquiry(e.target.value)} placeholder={t.inquiryContent}
              onKeyDown={e => e.key === "Enter" && addInquiry()}
              style={{ flex: 1, padding: "6px 10px", borderRadius: 7, border: "1px solid #fcd34d", fontSize: 12, boxSizing: "border-box" }} />
            <button onClick={addInquiry} style={{ fontSize: 12, padding: "6px 14px", borderRadius: 7, border: "none", background: "#f59e0b", color: "#fff", cursor: "pointer", whiteSpace: "nowrap" }}>{t.addInquiry}</button>
          </div>
          <div style={{ fontSize: 10, color: "#b45309", marginTop: 6 }}>※ 添加问询后状态自动变为「{t.qsInq}」/ 문의 추가 시 상태가 자동으로 「{t.qsInq}」로 변경됩니다</div>
        </div>

        {/* 产品列表 */}
        {form.items.map((item, idx) => {
          const { sel, purchaseKRW, profit, profitRate } = calcItemMoney(item);
          return (
            <div key={item.id} style={{ border: "1.5px solid #e2e8f0", borderRadius: 12, padding: 16, marginBottom: 16, background: "#fdfdfd" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <span style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontWeight: 600, fontSize: 14, color: "#1e293b" }}>📦 {t.items} {idx + 1}</span>
                  <AttachBtn attachments={item.attachments || []} onChange={atts => setItem(idx, { ...item, attachments: atts })} lang={lang} label={item.productName} />
                </span>
                {form.items.length > 1 && <button onClick={() => delItem(idx)} style={{ border: "none", background: "none", color: "#dc2626", cursor: "pointer", fontSize: 12 }}>✕ {t.delete}</button>}
              </div>

              {/* 调查书字段 */}
              <div style={{ display: "grid", gridTemplateColumns: "2fr 2fr 1.5fr 1fr", gap: "10px 14px" }}>
                <div>
                  <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{t.product}</label>
                  <input value={item.productName} onChange={e => setItem(idx, { ...item, productName: e.target.value })}
                    style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{t.detailQty}</label>
                  <input value={item.detailQty} onChange={e => setItem(idx, { ...item, detailQty: e.target.value })}
                    style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{t.colorSize}</label>
                  <input value={item.colorSize} onChange={e => setItem(idx, { ...item, colorSize: e.target.value })}
                    style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>{t.sampleNeeded}</label>
                  <select value={item.sampleNeeded ? "1" : "0"} onChange={e => setItem(idx, { ...item, sampleNeeded: e.target.value === "1" })}
                    style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12 }}>
                    <option value="0">{t.no}</option><option value="1">{t.yes}</option>
                  </select>
                </div>
              </div>
              <div style={{ marginTop: 10 }}>
                <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 3 }}>
                  {t.customerReq} <AITranslateBtn text={item.customerReq} lang={lang} fieldLabel={t.customerReq} />
                </label>
                <textarea value={item.customerReq} onChange={e => setItem(idx, { ...item, customerReq: e.target.value })} rows={2}
                  style={{ width: "100%", padding: "6px 9px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 12, resize: "vertical", boxSizing: "border-box" }} />
              </div>

              {/* 比价表 */}
              <div style={{ fontSize: 12, fontWeight: 600, color: "#475569", marginTop: 14 }}>🔍 {t.compareInfo} ({item.sources.length})</div>
              <SourceTable item={item} lang={lang} perm={perm} onChange={it => setItem(idx, it)} />

              {/* 该产品金额汇总 */}
              <div style={{ display: "flex", gap: 24, marginTop: 12, padding: "10px 14px", background: "#f0f9ff", borderRadius: 8, fontSize: 12, alignItems: "center", flexWrap: "wrap" }}>
                <span><span style={{ color: "#64748b" }}>{t.selected}: </span>
                  {sel ? <b>{sel.factory || sel.url?.slice(0, 30) || "—"}</b> : <span style={{ color: "#d97706" }}>⚠ {t.selectFirst}</span>}
                </span>
                <span><span style={{ color: "#64748b" }}>{t.landedCost}: </span><b style={{ color: "#0369a1" }}>₩{purchaseKRW.toLocaleString()}</b></span>
                <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ color: "#64748b" }}>{t.salePrice}: </span>
                  {perm.viewSale
                    ? <input type="number" value={item.salePrice || ""} disabled={!perm.editSale}
                        onChange={e => setItem(idx, { ...item, salePrice: e.target.value })}
                        style={{ width: 120, padding: "5px 8px", borderRadius: 6, border: "1px solid #d1d5db", fontSize: 12, background: perm.editSale ? "#fff" : "#f3f4f6" }} />
                    : <Lock lang={lang} />}
                </span>
                <span><span style={{ color: "#64748b" }}>{t.profit}: </span>
                  {perm.viewProfit
                    ? <b style={{ color: profit >= 0 ? "#166534" : "#dc2626" }}>₩{profit.toLocaleString()} ({profitRate.toFixed(1)}%)</b>
                    : <Lock lang={lang} />}
                </span>
              </div>
            </div>
          );
        })}

        <button onClick={addItem} style={{ fontSize: 13, padding: "8px 18px", borderRadius: 8, border: "1.5px dashed #93c5fd", background: "#f8fafc", color: "#1d4ed8", cursor: "pointer", width: "100%" }}>
          {t.addItem}
        </button>
        </fieldset>

        <div style={{ display: "flex", gap: 10, marginTop: 20, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={{ padding: "8px 20px", borderRadius: 8, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{locked ? t.close : t.cancel}</button>
          {!locked && <button onClick={() => onSave(form)} style={{ padding: "8px 20px", borderRadius: 8, border: "none", background: "#1a56db", color: "#fff", cursor: "pointer" }}>{t.save}</button>}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// 下单卡片 & 弹窗（沿用 v2，金额含内陆费）
// ============================================================

// ============================================================
// SMS 通知组件 — 每步状态变化时可发送短信给客户
// 演示模式: 显示预览+弹窗确认，点发送后模拟成功
// 正式系统: 将 sendSMS() 替换为 POST 到后端 /api/sms
//   后端再调用 SENS / 솔라피 / CoolSMS 等 API（密钥不可暴露在前端）
// ============================================================
function SmsButton({ order, lang, newStatus }) {
  const t = T[lang];
  const [open, setOpen] = useState(false);
  const [phone, setPhone] = useState(order.customerPhone || "");
  const [sent, setSent] = useState(false);
  const statusLabel = T[lang][newStatus] || newStatus;
  const msg = smsTemplate(order.id, order.product, statusLabel, order.deliveryDate);

  function handleSend() {
    // ▸ 正式系统替换为:
    // await fetch("/api/sms", { method:"POST", body: JSON.stringify({ to: phone, text: msg }) });
    console.log("[SMS DEMO]", { to: phone, text: msg });
    setSent(true);
    setTimeout(() => { setSent(false); setOpen(false); }, 2000);
  }

  return (
    <>
      <button onClick={() => setOpen(true)}
        style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #a7f3d0", background: "#ecfdf5", color: "#047857", cursor: "pointer", whiteSpace: "nowrap" }}>
        📱 {t.smsBtn.replace("📱 ", "")}
      </button>
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1500, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={() => setOpen(false)}>
          <div style={{ background: "#fff", borderRadius: 16, padding: 26, width: 400, boxShadow: "0 12px 40px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 14 }}>📱 {t.smsTitle}</div>

            {/* 手机号输入 */}
            <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.smsPhone}</label>
            <input value={phone} onChange={e => setPhone(e.target.value)} placeholder="010-XXXX-XXXX"
              style={{ width: "100%", padding: "8px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box", marginBottom: 14 }} />

            {/* 短信内容预览 */}
            <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: 10, padding: "12px 14px" }}>
              <div style={{ fontSize: 11, color: "#047857", fontWeight: 600, marginBottom: 6 }}>📨 {t.smsTitle}</div>
              <pre style={{ fontSize: 12, color: "#1e293b", margin: 0, whiteSpace: "pre-wrap", lineHeight: 1.6, fontFamily: "inherit" }}>{msg}</pre>
            </div>

            <div style={{ fontSize: 10, color: "#94a3b8", marginTop: 8, lineHeight: 1.5 }}>{t.smsPending}</div>

            <div style={{ display: "flex", gap: 10, marginTop: 16, justifyContent: "flex-end" }}>
              <button onClick={() => setOpen(false)} style={{ padding: "8px 18px", borderRadius: 8, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{t.smsCancel}</button>
              <button onClick={handleSend} disabled={!phone.trim() || sent}
                style={{ padding: "8px 18px", borderRadius: 8, border: "none", background: sent ? "#10b981" : "#047857", color: "#fff", fontWeight: 600, cursor: phone.trim() ? "pointer" : "not-allowed" }}>
                {sent ? t.smsSent : t.smsSend}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

function OrderCard({ o, lang, perm, user, onEdit, onDelete, onAdvance }) {
  const t = T[lang];
  const editable = canTouchCase(perm, user, o);
  const days = getDaysLeft(o.deliveryDate);
  const u = urgencyOf(days);
  const borderColor = u === "overdue" ? "#fecaca" : u === "urgent" ? "#fde68a" : "#e5e7eb";
  const { purchaseKRW, profit, profitRate } = calcOrderMoney(o);
  return (
    <div style={{ background: "#fff", border: `1.5px solid ${borderColor}`, borderRadius: 14, padding: "16px 20px", marginBottom: 14, boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <span style={{ fontSize: 11, color: "#9ca3af", fontFamily: "monospace" }}>{o.id}{o.linkedQuote ? ` ← ${o.linkedQuote}` : ""}</span>
          <div style={{ fontWeight: 600, fontSize: 15, marginTop: 2 }}>{o.customer}</div>
          <div style={{ fontSize: 13, color: "#374151", marginTop: 2 }}>{o.product}</div>
          <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 2 }}>{o.spec}</div>
        </div>
        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6 }}>
          <StatusBadge status={o.status} lang={lang} />
          <DaysBadge days={days} lang={lang} />
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "5px 16px", marginTop: 10, padding: "10px 12px", background: "#f8fafc", borderRadius: 8, fontSize: 12 }}>
        <div><span style={{ color: "#9ca3af" }}>{t.purchaseAmount}: </span>
          {perm.viewPurchase ? <b>{CURRENCY_SYMBOL[o.purchaseCurrency]}{(parseFloat(o.purchaseAmount)||0).toLocaleString()}</b> : <Lock lang={lang} />}
        </div>
        <div><span style={{ color: "#9ca3af" }}>{t.exchangeRate}: </span>
          {perm.viewPurchase ? <span>{o.lockedRate}</span> : "—"}
        </div>
        <div><span style={{ color: "#9ca3af" }}>{t.landedCost}: </span>
          {perm.viewPurchase ? <b style={{ color: "#0369a1" }}>₩{purchaseKRW.toLocaleString()}</b> : <Lock lang={lang} />}
        </div>
        <div><span style={{ color: "#9ca3af" }}>{t.salePrice}: </span>
          {perm.viewSale ? <b>₩{(parseFloat(o.salePrice)||0).toLocaleString()}</b> : <Lock lang={lang} />}
        </div>
        <div><span style={{ color: "#9ca3af" }}>{t.profit}: </span>
          {perm.viewProfit ? <b style={{ color: profit >= 0 ? "#166534" : "#dc2626" }}>₩{profit.toLocaleString()}</b> : <Lock lang={lang} />}
        </div>
        <div><span style={{ color: "#9ca3af" }}>{t.profitRate}: </span>
          {perm.viewProfit ? <b style={{ color: profitRate >= 0 ? "#166534" : "#dc2626" }}>{profitRate.toFixed(1)}%</b> : <Lock lang={lang} />}
        </div>
      </div>
      <div style={{ display: "flex", gap: 14, marginTop: 8, fontSize: 12, color: "#6b7280", flexWrap: "wrap" }}>
        <span>👤 {o.salesRep || "—"}</span>
        <span>📝 {o.purchaseRep || "—"}</span>
        <span>🚚 {o.deliveryDate || "—"}</span>
        {o.paymentStatus && <span style={{ padding: "1px 8px", borderRadius: 5, background: o.paymentStatus === "입금완료" ? "#dcfce7" : "#fef9c3", color: o.paymentStatus === "입금완료" ? "#166534" : "#92400e", fontSize: 11 }}>{o.paymentStatus}</span>}
      </div>
      <ProgressBar status={o.status} />
      {/* 下一步状态 + SMS按钮 */}
      {editable && (() => {
        const idx = ORDER_STATUSES.indexOf(o.status);
        const nextStatus = idx >= 0 && idx < ORDER_STATUSES.length - 1 ? ORDER_STATUSES[idx + 1] : null;
        return nextStatus ? (
          <div style={{ display: "flex", gap: 8, marginTop: 10, padding: "8px 10px", background: "#f0fdf4", borderRadius: 8, alignItems: "center" }}>
            <span style={{ fontSize: 11, color: "#047857", flex: 1 }}>→ {t[nextStatus]}</span>
            <button onClick={() => onAdvance(o.id, nextStatus)}
              style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "none", background: "#047857", color: "#fff", cursor: "pointer", whiteSpace: "nowrap" }}>
              ✓ {t[nextStatus]}
            </button>
            <SmsButton order={o} lang={lang} newStatus={nextStatus} />
          </div>
        ) : null;
      })()}
      <div style={{ display: "flex", gap: 8, marginTop: 8, justifyContent: "flex-end", alignItems: "center" }}>
        {!editable && <span style={{ fontSize: 11, color: "#d97706", marginRight: "auto" }}>🔒 {t.readOnlyOwn}</span>}
        <button onClick={() => onEdit(o)} style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{editable ? t.edit : t.detail}</button>
        {perm.canDelete && (
          <button onClick={() => onDelete(o.id)} style={{ fontSize: 12, padding: "5px 12px", borderRadius: 7, border: "1px solid #fecaca", background: "#fff5f5", color: "#dc2626", cursor: "pointer" }}>{t.delete}</button>
        )}
      </div>
    </div>
  );
}

function OrderModal({ initial, lang, perm, user, onSave, onClose }) {
  const t = T[lang];
  const locked = initial ? !canTouchCase(perm, user, initial) : false;
  const [form, setForm] = useState(initial || {
    id: `O${new Date().getFullYear()}-${String(Math.floor(Math.random()*900)+100)}`,
    linkedQuote: "", customer: "", product: "", spec: "", qty: "",
    purchaseCurrency: "CNY", purchaseAmount: "", lockedRate: LIVE_RATES.CNY, inlandFee: "",
    salePrice: "", supplier: "", salesRep: user?.name || "", purchaseRep: "", status: "os1", paymentStatus: "미결제",
    factoryDelivery: "", estShipDate: "", warehouseDate: "", loadingDate: "", shipInfo: "",
    arrivalDate: "", customsClear: "", deliveryDate: "", logistics: "", customerNote: ""
  });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const { purchaseKRW, profit, profitRate } = calcOrderMoney(form);
  const dateKeys = ["factoryDelivery","estShipDate","warehouseDate","loadingDate","arrivalDate","deliveryDate"];

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 999, display: "flex", alignItems: "center", justifyContent: "center" }} onClick={onClose}>
      <div style={{ background: "#fff", borderRadius: 16, padding: 28, width: 620, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 16px 48px rgba(0,0,0,0.18)" }} onClick={e => e.stopPropagation()}>
        <div style={{ fontWeight: 700, fontSize: 17, marginBottom: 18 }}>{initial ? (locked ? t.detail : t.edit) : t.newOrder} — {form.id}</div>
        {locked && (
          <div style={{ background: "#fffbeb", border: "1px solid #fcd34d", borderRadius: 9, padding: "8px 14px", fontSize: 12, color: "#b45309", marginBottom: 14 }}>
            🔒 {t.readOnlyOwn}
          </div>
        )}
        <fieldset disabled={locked} style={{ border: "none", padding: 0, margin: 0 }}>
        <div style={{ fontWeight: 500, fontSize: 13, color: "#6b7280", marginBottom: 10, borderBottom: "1px solid #f3f4f6", paddingBottom: 6 }}>{t.basicInfo}</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 16px" }}>
          {[["linkedQuote", t.linkedQuote], ["customer", t.customer], ["product", t.product], ["spec", t.spec], ["qty", t.qty], ["supplier", t.supplier], ["salesRep", t.salesRep], ["purchaseRep", t.purchaseRep], ["paymentStatus", t.paymentStatus]].map(([k, label]) => (
            <div key={k}>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{label}</label>
              <input type={k === "qty" ? "number" : "text"} value={form[k] || ""} onChange={e => set(k, e.target.value)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
            </div>
          ))}
          <div>
            <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.status}</label>
            <select value={form.status} onChange={e => set("status", e.target.value)} style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13 }}>
              {ORDER_STATUSES.map(s => <option key={s} value={s}>{t[s]}</option>)}
            </select>
          </div>
        </div>

        <div style={{ fontWeight: 500, fontSize: 13, color: "#6b7280", margin: "16px 0 10px", borderBottom: "1px solid #f3f4f6", paddingBottom: 6 }}>💰 {t.moneyInfo}</div>
        {perm.viewPurchase ? (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: "12px 14px" }}>
            <div>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.purchaseCurrency}</label>
              <select value={form.purchaseCurrency} disabled={!perm.editPurchase}
                onChange={e => { set("purchaseCurrency", e.target.value); set("lockedRate", LIVE_RATES[e.target.value] || 1); }}
                style={{ width: "100%", padding: "7px 8px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13 }}>
                {CURRENCIES.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.purchaseAmount}</label>
              <input type="number" value={form.purchaseAmount || ""} disabled={!perm.editPurchase} onChange={e => set("purchaseAmount", e.target.value)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
            </div>
            <div>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.inlandFee}</label>
              <input type="number" value={form.inlandFee || ""} disabled={!perm.editPurchase} onChange={e => set("inlandFee", e.target.value)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
            </div>
            <div>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.lockedRate}</label>
              <input type="number" step="0.01" value={form.lockedRate || ""} disabled={!perm.editRate} onChange={e => set("lockedRate", e.target.value)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1.5px solid #93c5fd", fontSize: 13, boxSizing: "border-box" }} />
            </div>
            <div style={{ gridColumn: "span 2" }}>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.landedCost}</label>
              <div style={{ padding: "7px 10px", borderRadius: 7, border: "1px solid #e5e7eb", fontSize: 13, background: "#f0f9ff", color: "#0369a1", fontWeight: 600 }}>
                ₩{purchaseKRW.toLocaleString()}
              </div>
            </div>
            {perm.viewSale ? (
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.salePrice}</label>
                <input type="number" value={form.salePrice || ""} disabled={!perm.editSale} onChange={e => set("salePrice", e.target.value)}
                  style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box", background: perm.editSale ? "#fff" : "#f3f4f6" }} />
              </div>
            ) : (
              <div style={{ gridColumn: "span 2", padding: "7px 10px", borderRadius: 7, border: "1px dashed #e5e7eb", fontSize: 12, color: "#d1d5db", fontStyle: "italic", alignSelf: "end" }}>🔒 {t.salePrice} — {t.noPermission}</div>
            )}
            {perm.viewProfit ? (
              <>
                <div style={{ gridColumn: "span 2" }}>
                  <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.profit} / {t.profitRate}</label>
                  <div style={{ padding: "7px 10px", borderRadius: 7, border: "1px solid #e5e7eb", fontSize: 13, background: "#f9fafb", fontWeight: 600, color: profit >= 0 ? "#166534" : "#dc2626" }}>
                    ₩{profit.toLocaleString()} ({profitRate.toFixed(1)}%)
                  </div>
                </div>
              </>
            ) : (
              <div style={{ gridColumn: "span 2", padding: "7px 10px", borderRadius: 7, border: "1px dashed #e5e7eb", fontSize: 12, color: "#d1d5db", fontStyle: "italic", alignSelf: "end" }}>🔒 {t.profit} — {t.noPermission}</div>
            )}
          </div>
        ) : (
          <div style={{ padding: "10px 12px", background: "#f9fafb", borderRadius: 8, fontSize: 12, color: "#d1d5db", fontStyle: "italic" }}>🔒 {t.moneyInfo} — {t.noPermission}</div>
        )}

        <div style={{ fontWeight: 500, fontSize: 13, color: "#6b7280", margin: "16px 0 10px", borderBottom: "1px solid #f3f4f6", paddingBottom: 6 }}>🚚 {t.logisticsInfo}</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "12px 16px" }}>
          {[["factoryDelivery", t.factoryDelivery], ["estShipDate", t.estShipDate], ["warehouseDate", t.warehouseDate], ["loadingDate", t.loadingDate], ["shipInfo", t.shipInfo], ["arrivalDate", t.arrivalDate], ["customsClear", t.customsClear], ["deliveryDate", t.deliveryDate]].map(([k, label]) => (
            <div key={k}>
              <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>{label}</label>
              <input type={dateKeys.includes(k) ? "date" : "text"} value={form[k] || ""} onChange={e => set(k, e.target.value)}
                style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
            </div>
          ))}
        </div>

        {[["logistics", t.logistics], ["customerNote", t.customerNote]].map(([k, label]) => (
          <div key={k} style={{ marginTop: 12 }}>
            <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 4 }}>
              {label} <AITranslateBtn text={form[k]} lang={lang} fieldLabel={label} />
            </label>
            <textarea value={form[k] || ""} onChange={e => set(k, e.target.value)} rows={2}
              style={{ width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13, resize: "vertical", boxSizing: "border-box" }} />
          </div>
        ))}
        </fieldset>

        <div style={{ display: "flex", gap: 10, marginTop: 20, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={{ padding: "8px 20px", borderRadius: 8, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{locked ? t.close : t.cancel}</button>
          {!locked && <button onClick={() => onSave(form)} style={{ padding: "8px 20px", borderRadius: 8, border: "none", background: "#1a56db", color: "#fff", cursor: "pointer" }}>{t.save}</button>}
        </div>
      </div>
    </div>
  );
}

// ============================================================
// Dashboard
// ============================================================
function Dashboard({ quotes, orders, lang, perm, onNav }) {
  const t = T[lang];
  const all = [...quotes.map(q => ({ ...q, _d: q.estimatedDelivery, _p: q.items.map(i => i.productName).join("·") })), ...orders.map(o => ({ ...o, _d: o.deliveryDate, _p: o.product }))];
  const urgent = all.filter(c => { const d = getDaysLeft(c._d); return d !== null && d >= 0 && d <= 5; });
  const overdue = all.filter(c => { const d = getDaysLeft(c._d); return d !== null && d < 0; });
  const stats = [
    { label: t.totalQuotes, value: quotes.length, icon: "📋", color: "#eff6ff", tc: "#1a56db" },
    { label: t.totalOrders, value: orders.length, icon: "📦", color: "#f0fdf4", tc: "#166534" },
    { label: t.urgentCases, value: urgent.length, icon: "⚠️", color: "#fefce8", tc: "#92400e" },
    { label: t.overdueCases, value: overdue.length, icon: "🔴", color: "#fff1f2", tc: "#be123c" },
  ];
  const totalProfit = quotes.reduce((s, q) => s + q.items.reduce((si, it) => si + calcItemMoney(it).profit, 0), 0)
    + orders.reduce((s, o) => s + calcOrderMoney(o).profit, 0);
  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 18 }}>
        {stats.map(s => (
          <div key={s.label} style={{ background: s.color, borderRadius: 12, padding: "16px 18px", border: "1px solid rgba(0,0,0,0.06)" }}>
            <div style={{ fontSize: 20 }}>{s.icon}</div>
            <div style={{ fontSize: 26, fontWeight: 700, color: s.tc, marginTop: 4 }}>{s.value}</div>
            <div style={{ fontSize: 12, color: "#6b7280" }}>{s.label}</div>
          </div>
        ))}
      </div>
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 12, padding: "14px 18px", marginBottom: 18, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 13, color: "#6b7280" }}>💰 {t.profit} ({lang === "zh" ? "全部案件合计" : "전체 합계"})</span>
        {perm.viewProfit
          ? <span style={{ fontSize: 20, fontWeight: 700, color: totalProfit >= 0 ? "#166534" : "#dc2626" }}>₩{Math.round(totalProfit).toLocaleString()}</span>
          : <Lock lang={lang} />}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 18 }}>
        <div onClick={() => onNav("quote")} style={{ background: "#fff", border: "1.5px solid #bfdbfe", borderRadius: 14, padding: 18, cursor: "pointer" }}>
          <div style={{ fontWeight: 600, fontSize: 15, color: "#1a56db" }}>📋 {t.quoteSystem}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{quotes.length}건</div>
        </div>
        <div onClick={() => onNav("order")} style={{ background: "#fff", border: "1.5px solid #bbf7d0", borderRadius: 14, padding: 18, cursor: "pointer" }}>
          <div style={{ fontWeight: 600, fontSize: 15, color: "#166534" }}>📦 {t.orderSystem}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{orders.length}건</div>
        </div>
      </div>
      <div style={{ background: "#fff", borderRadius: 14, border: "1px solid #e5e7eb", padding: "16px 20px" }}>
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12, color: "#374151" }}>{t.recentUpdated}</div>
        {all.slice(0, 6).map(c => (
          <div key={c.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", borderBottom: "1px solid #f3f4f6" }}>
            <div style={{ minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
              <span style={{ fontSize: 11, color: "#9ca3af", fontFamily: "monospace", marginRight: 8 }}>{c.id}</span>
              <span style={{ fontSize: 13, fontWeight: 500 }}>{c.customer}</span>
              <span style={{ fontSize: 12, color: "#6b7280", marginLeft: 8 }}>{c._p}</span>
            </div>
            <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
              <StatusBadge status={c.status} lang={lang} />
              <DaysBadge days={getDaysLeft(c._d)} lang={lang} />
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

// ============================================================
// Excel 导出（按角色过滤；正式系统须在后端做过滤）
// ============================================================
function exportQuotesCSV(quotes, perm, lang) {
  const t = T[lang];
  const head = [t.caseNo, t.customer, t.product, t.factory, t.unitPrice, t.inlandFee, t.packing, t.packingFee, t.perBox, t.boxCount, t.cbm, t.weightKg, t.leadTime, t.sampleNote, t.landedCost];
  if (perm.viewSale) head.push(t.salePrice);
  if (perm.viewProfit) head.push(t.profit, t.profitRate);
  const rows = [];
  quotes.forEach(q => q.items.forEach(it => it.sources.forEach(s => {
    const krw = calcSourceKRW(s);
    const m = calcItemMoney(it);
    const r = [q.id, q.customer, it.productName, s.factory, s.unitPrice, s.inlandFee, s.packing, s.packingFee, s.perBox, s.boxCount, s.cbm, s.weightKg, s.leadTime, s.sampleNote, krw];
    if (perm.viewSale) r.push(s.selected ? it.salePrice : "");
    if (perm.viewProfit) { r.push(s.selected ? Math.round(m.profit) : ""); r.push(s.selected ? m.profitRate.toFixed(1) + "%" : ""); }
    rows.push(r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","));
  })));
  const csv = "\uFEFF" + [head.join(","), ...rows].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "quotes_export.csv";
  a.click();
}


// ============================================================
// 엑셀 내보내기 — SheetJS(xlsx) 라이브러리 사용
// 브라우저에서 직접 진짜 .xlsx 파일 생성 (서버 불필요)
// ============================================================

// ── 공통 스타일 헬퍼 ──────────────────────────────────────
function makeWb() { return XLSX.utils.book_new(); }
function addSheet(wb, data, name) {
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, name);
  return ws;
}
function setColWidths(ws, widths) {
  ws['!cols'] = widths.map(w => ({ wch: w }));
}
function saveWb(wb, filename) {
  XLSX.writeFile(wb, filename);
}

// ── 1. 발주 리스트 (주문 시스템) ────────────────────────────
// 시트1: 발주 리스트 전체 (한 행 = 한 주문)
// ──────────────────────────────────────────────────────────
function exportOrderListXLSX(orders, perm, lang) {
  const t = T[lang];
  const today = new Date().toISOString().slice(0, 10);
  const wb = makeWb();

  // 헤더 행
  const header = [
    t.orderNo, t.linkedQuote, t.customer,
    t.product, t.spec, t.qty,
    t.salesRep, t.purchaseRep, t.supplier,
    t.status, t.paymentStatus,
    t.factoryDelivery, t.estShipDate, t.warehouseDate,
    t.loadingDate, t.shipInfo, t.arrivalDate,
    t.customsClear, t.deliveryDate,
  ];
  if (perm.viewPurchase) {
    header.push(t.purchaseCurrency, t.purchaseAmount, t.inlandFee, t.exchangeRate, t.landedCost);
  }
  if (perm.viewSale) header.push(t.salePrice);
  if (perm.viewProfit) header.push(t.profit, t.profitRate);
  header.push(t.logistics, t.customerNote);

  const rows = [header];
  orders.forEach(o => {
    const { purchaseKRW, profit, profitRate } = calcOrderMoney(o);
    const row = [
      o.id, o.linkedQuote || '', o.customer,
      o.product, o.spec, o.qty,
      o.salesRep || '', o.purchaseRep || '', o.supplier || '',
      t[o.status] || o.status, o.paymentStatus || '',
      o.factoryDelivery || '', o.estShipDate || '', o.warehouseDate || '',
      o.loadingDate || '', o.shipInfo || '', o.arrivalDate || '',
      o.customsClear || '', o.deliveryDate || '',
    ];
    if (perm.viewPurchase) {
      row.push(o.purchaseCurrency, parseFloat(o.purchaseAmount)||0, parseFloat(o.inlandFee)||0, parseFloat(o.lockedRate)||1, purchaseKRW);
    }
    if (perm.viewSale) row.push(parseFloat(o.salePrice)||0);
    if (perm.viewProfit) { row.push(Math.round(profit)); row.push(parseFloat(profitRate.toFixed(1))); }
    row.push(o.logistics || '', o.customerNote || '');
    rows.push(row);
  });

  // 합계 행 (금액 권한 있을 때만)
  if (perm.viewSale || perm.viewProfit) {
    const sumRow = new Array(header.length).fill('');
    sumRow[0] = lang === 'zh' ? '合计' : '합계';
    if (perm.viewPurchase) {
      const landedIdx = header.indexOf(t.landedCost);
      if (landedIdx >= 0) sumRow[landedIdx] = orders.reduce((s,o) => s + calcOrderMoney(o).purchaseKRW, 0);
    }
    if (perm.viewSale) {
      const saleIdx = header.indexOf(t.salePrice);
      if (saleIdx >= 0) sumRow[saleIdx] = orders.reduce((s,o) => s + (parseFloat(o.salePrice)||0), 0);
    }
    if (perm.viewProfit) {
      const profitIdx = header.indexOf(t.profit);
      if (profitIdx >= 0) sumRow[profitIdx] = Math.round(orders.reduce((s,o) => s + calcOrderMoney(o).profit, 0));
    }
    rows.push(sumRow);
  }

  const ws = addSheet(wb, rows, t.sheetOrderList);
  setColWidths(ws, header.map((_, i) => i < 3 ? 14 : i < 9 ? 18 : 12));
  saveWb(wb, `발주리스트_${today}.xlsx`);
}

// ── 2. 견적서 (견적 시스템) ──────────────────────────────────
// 시트1: 견적 요약 (한 행 = 한 케이스)
// 시트2: 비교 견적 상세 (한 행 = 한 공급업체 견적 행)
// ──────────────────────────────────────────────────────────
function exportQuoteSheetXLSX(quotes, perm, lang) {
  const t = T[lang];
  const today = new Date().toISOString().slice(0, 10);
  const wb = makeWb();

  // ─ 시트1: 견적 요약 ─
  const sumHeader = [
    t.caseNo, t.category, t.customer, t.salesRep, t.purchaseRep,
    t.product, t.detailQty, t.status, t.quoteDate, t.estimatedDelivery,
    lang==='zh'?'선정 공장':'선정 공장', lang==='zh'?'선정 단가':'선정 단가',
    t.purchaseCurrency, t.exchangeRate, t.landedCost,
  ];
  if (perm.viewSale) sumHeader.push(t.salePrice);
  if (perm.viewProfit) sumHeader.push(t.profit, t.profitRate);
  sumHeader.push(t.customerReq, t.sampleNote, t.note);

  const sumRows = [sumHeader];
  quotes.forEach(q => {
    q.items.forEach(it => {
      const { sel, purchaseKRW, profit, profitRate } = calcItemMoney(it);
      const row = [
        q.id,
        q.category === 'machine' ? (lang==='zh'?'机器类':'기계류') : (lang==='zh'?'日常生活用品':'생활용품'),
        q.customer, q.salesRep||'', q.purchaseRep||'',
        it.productName, it.detailQty||'',
        t[q.status]||q.status, q.quoteDate||'', q.estimatedDelivery||'',
        sel ? (sel.factory||sel.url||'') : (lang==='zh'?'未选定':'미선정'),
        sel ? (parseFloat(sel.unitPrice)||0) : '',
        sel ? sel.currency : '',
        sel ? (parseFloat(sel.lockedRate)||1) : '',
        sel ? purchaseKRW : '',
      ];
      if (perm.viewSale) row.push(parseFloat(it.salePrice)||0);
      if (perm.viewProfit) { row.push(Math.round(profit)); row.push(parseFloat(profitRate.toFixed(1))); }
      row.push(it.customerReq||'', sel?.sampleNote||'', '');
      sumRows.push(row);
    });
  });

  const ws1 = addSheet(wb, sumRows, t.sheetQuoteSummary);
  setColWidths(ws1, sumHeader.map((_,i) => i===0?14 : i<5?12 : i<7?20 : 12));

  // ─ 시트2: 비교 견적 상세 ─
  const detHeader = [
    t.caseNo, t.customer, t.product,
    t.url, t.factory, t.unitPrice, t.purchaseCurrency, t.qty,
    t.inlandFee, t.packing, t.packingFee,
    t.perBox, t.boxCount, t.cbm, t.weightKg,
    t.origin, t.leadTime, t.sampleNote, t.note,
    t.exchangeRate, t.landedCost,
    lang==='zh'?'선정여부':'선정여부',
  ];
  const detRows = [detHeader];
  quotes.forEach(q => {
    q.items.forEach(it => {
      it.sources.forEach(s => {
        const krw = calcSourceKRW(s);
        detRows.push([
          q.id, q.customer, it.productName,
          s.url||'', s.factory||'',
          parseFloat(s.unitPrice)||0, s.currency||'CNY', parseFloat(s.qty)||1,
          parseFloat(s.inlandFee)||0, s.packing||'', parseFloat(s.packingFee)||0,
          parseFloat(s.perBox)||'', parseFloat(s.boxCount)||'',
          parseFloat(s.cbm)||'', parseFloat(s.weightKg)||'',
          s.origin||'', s.leadTime||'', s.sampleNote||'', s.note||'',
          parseFloat(s.lockedRate)||1, s.noQuote ? '' : krw,
          s.noQuote ? (lang==='zh'?'거절':'거절') : s.selected ? (lang==='zh'?'✓ 선정':'✓ 선정') : '',
        ]);
      });
    });
  });

  const ws2 = addSheet(wb, detRows, t.sheetPriceCompare);
  setColWidths(ws2, detHeader.map((_,i) => i===3?40 : i<5?16 : 10));

  saveWb(wb, `견적서_${today}.xlsx`);
}


// ============================================================
// 내보내기 선택 모달 — 발주 리스트 / 견적서 선택
// ============================================================
function ExportModal({ quotes, orders, filteredQuotes, filteredOrders, perm, lang, onClose }) {
  const t = T[lang];
  const [scope, setScope] = useState("filtered");
  const [exporting, setExporting] = useState(false);

  function doExport(type) {
    setExporting(true);
    setTimeout(() => {
      const q = scope === "all" ? quotes : filteredQuotes;
      const o = scope === "all" ? orders : filteredOrders;
      if (type === "order") exportOrderListXLSX(o, perm, lang);
      if (type === "quote") exportQuoteSheetXLSX(q, perm, lang);
      setExporting(false);
      onClose();
    }, 100);
  }

  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.45)", zIndex:1500, display:"flex", alignItems:"center", justifyContent:"center" }} onClick={onClose}>
      <div style={{ background:"#fff", borderRadius:16, padding:28, width:420, boxShadow:"0 12px 40px rgba(0,0,0,0.2)" }} onClick={e=>e.stopPropagation()}>
        <div style={{ fontWeight:700, fontSize:16, marginBottom:18 }}>📥 {t.exportChoose}</div>

        {/* 범위 선택 */}
        <div style={{ display:"flex", gap:10, marginBottom:20 }}>
          {[["filtered", t.exportFiltered], ["all", t.exportAll]].map(([k,label]) => (
            <button key={k} onClick={() => setScope(k)}
              style={{ flex:1, padding:"9px 0", borderRadius:9, border: scope===k?"2px solid #1a56db":"1px solid #d1d5db", background: scope===k?"#eff6ff":"#f9fafb", color: scope===k?"#1a56db":"#6b7280", fontWeight: scope===k?700:400, fontSize:13, cursor:"pointer" }}>
              {label}
            </button>
          ))}
        </div>

        {/* 발주 리스트 */}
        <div style={{ border:"1.5px solid #bbf7d0", borderRadius:12, padding:"16px 18px", marginBottom:14, cursor:exporting?"not-allowed":"pointer" }}
          onClick={() => !exporting && doExport("order")}>
          <div style={{ fontWeight:700, fontSize:14, color:"#047857" }}>📊 {t.exportOrderList}</div>
          <div style={{ fontSize:12, color:"#6b7280", marginTop:5, lineHeight:1.6 }}>
            {lang==='zh'
              ? `주문번호·고객명·제품·납기·진행상태·${perm.viewPurchase?'매입가·':''}${perm.viewSale?'판매가·':''}물류 일정을 한 줄씩 정리`
              : `주문번호·고객명·제품·납기·진행상태·${perm.viewPurchase?'매입가·':''}${perm.viewSale?'판매가·':''}물류 일정을 한 줄씩 정리`}
          </div>
          <div style={{ fontSize:11, color:"#9ca3af", marginTop:4 }}>
            {lang==='zh'?'대상:':'대상:'} {(scope==="all"?orders:filteredOrders).length}건 → <b>발주리스트_{new Date().toISOString().slice(0,10)}.xlsx</b>
          </div>
        </div>

        {/* 견적서 */}
        <div style={{ border:"1.5px solid #bfdbfe", borderRadius:12, padding:"16px 18px", cursor:exporting?"not-allowed":"pointer" }}
          onClick={() => !exporting && doExport("quote")}>
          <div style={{ fontWeight:700, fontSize:14, color:"#1a56db" }}>📋 {t.exportQuoteSheet}</div>
          <div style={{ fontSize:12, color:"#6b7280", marginTop:5, lineHeight:1.6 }}>
            {lang==='zh'
              ? '시트①: 견적 요약 (선정 공급업체+금액) / 시트②: 비교 견적 상세 전체'
              : '시트①: 견적 요약 (선정 공급업체+금액) / 시트②: 비교 견적 상세 전체'}
          </div>
          <div style={{ fontSize:11, color:"#9ca3af", marginTop:4 }}>
            {lang==='zh'?'대상:':'대상:'} {(scope==="all"?quotes:filteredQuotes).length}건 → <b>견적서_{new Date().toISOString().slice(0,10)}.xlsx</b>
          </div>
        </div>

        <div style={{ fontSize:10, color:"#94a3b8", marginTop:14, lineHeight:1.5 }}>
          ※ {lang==='zh'
            ? '금액 항목은 현재 역할 권한에 따라 자동 필터링됩니다'
            : '금액 항목은 현재 역할 권한에 따라 자동 필터링됩니다'}
        </div>
        <div style={{ display:"flex", justifyContent:"flex-end", marginTop:16 }}>
          <button onClick={onClose} style={{ padding:"8px 20px", borderRadius:8, border:"1px solid #d1d5db", background:"#f9fafb", cursor:"pointer" }}>{t.close}</button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// 登录页 — 输入姓名或ID
// ============================================================
function LoginScreen({ users, lang, setLang, onLogin }) {
  const t = T[lang];
  const [idInput, setIdInput] = useState("");
  const [pwInput, setPwInput] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [error, setError] = useState("");
  function handleLogin() {
    const id = idInput.trim();
    const pw = pwInput;
    if (!id || !pw) return;
    const u = users.find(u => u.id.toLowerCase() === id.toLowerCase() || u.name === id);
    if (!u) { setError(t.userNotFound); return; }
    if (u.pw !== pw) { setError(t.wrongPw); return; }
    onLogin(u);
  }
  const clearErr = () => setError("");
  return (
    <div style={{ minHeight: "100vh", background: "#0f172a", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Noto Sans KR', 'PingFang SC', sans-serif" }}>
      <div style={{ background: "#fff", borderRadius: 18, padding: "40px 44px", width: 400, boxShadow: "0 24px 64px rgba(0,0,0,0.35)" }}>
        <div style={{ textAlign: "center", marginBottom: 26 }}>
          <div style={{ fontSize: 32 }}>🌏</div>
          <div style={{ fontWeight: 700, fontSize: 18, marginTop: 8 }}>{t.appTitle}</div>
          <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 4 }}>{t.loginTitle}</div>
        </div>
        <input value={idInput} autoFocus
          onChange={e => { setIdInput(e.target.value); clearErr(); }}
          onKeyDown={e => e.key === "Enter" && document.getElementById("pw-input").focus()}
          placeholder={t.enterNameOrId}
          style={{ width: "100%", padding: "11px 14px", borderRadius: 9, border: error ? "1.5px solid #f87171" : "1.5px solid #d1d5db", fontSize: 14, boxSizing: "border-box", marginBottom: 10, outline: "none" }} />
        <div style={{ position: "relative" }}>
          <input id="pw-input" type={showPw ? "text" : "password"} value={pwInput}
            onChange={e => { setPwInput(e.target.value); clearErr(); }}
            onKeyDown={e => e.key === "Enter" && handleLogin()}
            placeholder={t.enterPw}
            style={{ width: "100%", padding: "11px 44px 11px 14px", borderRadius: 9, border: error ? "1.5px solid #f87171" : "1.5px solid #d1d5db", fontSize: 14, boxSizing: "border-box", outline: "none" }} />
          <button onClick={() => setShowPw(s => !s)} style={{ position: "absolute", right: 12, top: "50%", transform: "translateY(-50%)", border: "none", background: "none", color: "#9ca3af", cursor: "pointer", fontSize: 11 }}>
            {showPw ? t.hidePw : t.showPw}
          </button>
        </div>
        {error && <div style={{ fontSize: 12, color: "#dc2626", marginTop: 8 }}>{error}</div>}
        <button onClick={handleLogin}
          style={{ width: "100%", marginTop: 14, padding: "12px 0", borderRadius: 9, border: "none", background: "#1a56db", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer" }}>
          {t.loginBtn}
        </button>
        <div style={{ marginTop: 22, padding: "12px 14px", background: "#f8fafc", borderRadius: 9, fontSize: 11, color: "#64748b", lineHeight: 1.9 }}>
          <b>{t.demoAccounts}:</b>
          {MOCK_USERS.map(u => (
            <div key={u.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 3, cursor: "pointer" }}
              onClick={() => { setIdInput(u.id); setPwInput(u.pw); clearErr(); }}>
              <span>
                <code style={{ background: "#e2e8f0", padding: "0 5px", borderRadius: 4 }}>{u.id}</code>
                <span style={{ marginLeft: 6 }}>{u.name}</span>
              </span>
              <span style={{ color: "#94a3b8" }}>{PERMISSIONS[u.role].label[lang]}</span>
            </div>
          ))}
          <div style={{ marginTop: 6, color: "#94a3b8", fontSize: 10 }}>↑ 클릭하면 자동 입력 / 点击自动填入</div>
        </div>
        <button onClick={() => setLang(l => l === "zh" ? "ko" : "zh")}
          style={{ width: "100%", marginTop: 12, padding: "8px 0", borderRadius: 8, border: "1px solid #e5e7eb", background: "#fff", color: "#6b7280", fontSize: 12, cursor: "pointer" }}>
          🌐 {t.lang}
        </button>
      </div>
    </div>
  );
}

// ============================================================
// 员工管理页（仅管理员可见）— 添加员工、分组、删除
// ============================================================
function StaffPage({ users, setUsers, currentUser, lang }) {
  const t = T[lang];
  const [newUser, setNewUser] = useState({ id: "", name: "", role: "sales", pw: "" });
  const [addError, setAddError] = useState("");
  const [editTarget, setEditTarget] = useState(null);
  const [editName, setEditName] = useState("");
  const [editPw, setEditPw] = useState("");
  const [editPwConfirm, setEditPwConfirm] = useState("");
  const [editShowPw, setEditShowPw] = useState(false);
  const [editError, setEditError] = useState("");

  const ROLE_COLORS = {
    admin:      ["#ede9fe", "#6d28d9"],
    sales:      ["#e8f0fe", "#1a56db"],
    purchasing: ["#e6f4ea", "#1a7f37"],
    parttime:   ["#fef9c3", "#92400e"],
  };

  function addStaff() {
    const id = newUser.id.trim(), name = newUser.name.trim(), pw = newUser.pw.trim();
    if (!id || !name || !pw) return;
    if (users.some(u => u.id.toLowerCase() === id.toLowerCase())) { setAddError(t.staffExists); return; }
    setUsers(us => [...us, { id, name, role: newUser.role, pw }]);
    setNewUser({ id: "", name: "", role: "sales", pw: "" });
    setAddError("");
  }
  function removeStaff(id) {
    if (id === currentUser.id) { alert(t.deleteSelf); return; }
    if (confirm("삭제하시겠습니까? / 确定删除？")) setUsers(us => us.filter(u => u.id !== id));
  }
  function changeRole(id, role) {
    setUsers(us => us.map(u => u.id === id ? { ...u, role } : u));
  }
  function openEdit(u) {
    setEditTarget(u);
    setEditName(u.name);
    setEditPw(""); setEditPwConfirm(""); setEditShowPw(false); setEditError("");
  }
  function saveEdit() {
    const name = editName.trim();
    if (!name) { setEditError(lang === "zh" ? "姓名不能为空" : "이름을 입력하세요"); return; }
    if (editPw && editPw !== editPwConfirm) {
      setEditError(lang === "zh" ? "两次密码不一致" : "비밀번호가 일치하지 않습니다"); return;
    }
    setUsers(us => us.map(u => u.id !== editTarget.id ? u : { ...u, name, pw: editPw.trim() || u.pw }));
    setEditTarget(null);
  }

  return (
    <div style={{ maxWidth: 780 }}>
      <div style={{ fontSize: 12, color: "#9ca3af", marginBottom: 16 }}>{t.staffHint}</div>

      {/* ── 직원 추가 ── */}
      <div style={{ background: "#fff", border: "1.5px solid #bfdbfe", borderRadius: 12, padding: "18px 20px", marginBottom: 20 }}>
        <div style={{ fontWeight: 600, fontSize: 13, color: "#1a56db", marginBottom: 12 }}>+ {t.addStaff}</div>
        <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
          {[["id", t.userId, "kim05", 110], ["name", t.userName, "홍길동 / 张三", 140], ["pw", t.userPw, "••••••", 110]].map(([k, label, ph, w]) => (
            <div key={k}>
              <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 4 }}>{label}</label>
              <input type={k === "pw" ? "password" : "text"} value={newUser[k]}
                onChange={e => { setNewUser(n => ({ ...n, [k]: e.target.value })); setAddError(""); }}
                onKeyDown={e => e.key === "Enter" && addStaff()}
                placeholder={ph} style={{ width: w, padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13 }} />
            </div>
          ))}
          <div>
            <label style={{ fontSize: 11, color: "#6b7280", display: "block", marginBottom: 4 }}>{t.userRole}</label>
            <select value={newUser.role} onChange={e => setNewUser(n => ({ ...n, role: e.target.value }))}
              style={{ padding: "7px 10px", borderRadius: 7, border: "1px solid #d1d5db", fontSize: 13 }}>
              {Object.entries(PERMISSIONS).map(([k, p]) => <option key={k} value={k}>{p.label[lang]}</option>)}
            </select>
          </div>
          <button onClick={addStaff}
            style={{ padding: "8px 20px", borderRadius: 8, border: "none", background: "#1a56db", color: "#fff", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
            + {t.addStaff}
          </button>
        </div>
        {addError && <div style={{ fontSize: 12, color: "#dc2626", marginTop: 8 }}>{addError}</div>}
      </div>

      {/* ── 직원 목록 ── */}
      <div style={{ background: "#fff", border: "1px solid #e5e7eb", borderRadius: 12, overflow: "hidden" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#f8fafc" }}>
              {[t.userId, t.userName, t.userRole, t.userPw, ""].map((h, i) => (
                <th key={i} style={{ padding: "10px 16px", textAlign: "left", color: "#475569", fontWeight: 600, width: i === 4 ? 110 : "auto" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {users.map(u => {
              const [bg, color] = ROLE_COLORS[u.role] || ["#f1f5f9", "#475569"];
              const isSelf = u.id === currentUser.id;
              return (
                <tr key={u.id} style={{ borderTop: "1px solid #f3f4f6", background: isSelf ? "#fefce8" : "#fff" }}>
                  <td style={{ padding: "10px 16px", fontFamily: "monospace", color: "#64748b", fontSize: 12 }}>
                    {u.id} {isSelf && <span style={{ fontSize: 10, color: "#f59e0b", fontWeight: 700 }}>ME</span>}
                  </td>
                  <td style={{ padding: "10px 16px", fontWeight: 600 }}>{u.name}</td>
                  <td style={{ padding: "10px 16px" }}>
                    <select value={u.role} onChange={e => changeRole(u.id, e.target.value)}
                      style={{ padding: "5px 10px", borderRadius: 6, border: "none", background: bg, color, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                      {Object.entries(PERMISSIONS).map(([k, p]) => <option key={k} value={k}>{p.label[lang]}</option>)}
                    </select>
                  </td>
                  <td style={{ padding: "10px 16px", color: "#94a3b8", letterSpacing: 3, fontSize: 13 }}>
                    {"●".repeat(Math.min(u.pw?.length || 6, 8))}
                  </td>
                  <td style={{ padding: "10px 12px", textAlign: "center", whiteSpace: "nowrap" }}>
                    <button onClick={() => openEdit(u)}
                      style={{ fontSize: 12, padding: "4px 12px", borderRadius: 6, border: "1px solid #93c5fd", background: "#eff6ff", color: "#1d4ed8", cursor: "pointer", marginRight: 6 }}>
                      ✏️ {lang === "zh" ? "编辑" : "편집"}
                    </button>
                    <button onClick={() => removeStaff(u.id)} disabled={isSelf}
                      style={{ border: "none", background: "none", color: isSelf ? "#d1d5db" : "#dc2626", cursor: isSelf ? "not-allowed" : "pointer", fontSize: 14 }}>✕</button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* ── 편집 모달 ── */}
      {editTarget && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1500, display: "flex", alignItems: "center", justifyContent: "center" }}
          onClick={() => setEditTarget(null)}>
          <div style={{ background: "#fff", borderRadius: 16, padding: 28, width: 420, boxShadow: "0 12px 40px rgba(0,0,0,0.2)" }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 700, fontSize: 16, marginBottom: 20 }}>
              ✏️ {lang === "zh" ? "编辑员工信息" : "직원 정보 편집"}
              <span style={{ fontFamily: "monospace", color: "#64748b", fontSize: 13, marginLeft: 8 }}>{editTarget.id}</span>
            </div>

            <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 6 }}>{t.userName}</label>
            <input value={editName} onChange={e => { setEditName(e.target.value); setEditError(""); }}
              style={{ width: "100%", padding: "9px 12px", borderRadius: 8, border: "1.5px solid #93c5fd", fontSize: 14, boxSizing: "border-box", fontWeight: 500, marginBottom: 16 }} />

            <label style={{ fontSize: 12, color: "#6b7280", display: "block", marginBottom: 6 }}>
              {t.userPw}
              <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 6 }}>({lang === "zh" ? "留空则不修改" : "빈칸이면 변경 안 함"})</span>
            </label>
            <div style={{ position: "relative", marginBottom: 10 }}>
              <input type={editShowPw ? "text" : "password"} value={editPw}
                onChange={e => { setEditPw(e.target.value); setEditError(""); }}
                placeholder={lang === "zh" ? "新密码" : "새 비밀번호"}
                style={{ width: "100%", padding: "9px 44px 9px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 13, boxSizing: "border-box" }} />
              <button onClick={() => setEditShowPw(s => !s)}
                style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", border: "none", background: "none", color: "#9ca3af", cursor: "pointer", fontSize: 11 }}>
                {editShowPw ? t.hidePw : t.showPw}
              </button>
            </div>

            {editPw && (
              <div style={{ marginBottom: 10 }}>
                <input type={editShowPw ? "text" : "password"} value={editPwConfirm}
                  onChange={e => { setEditPwConfirm(e.target.value); setEditError(""); }}
                  placeholder={lang === "zh" ? "确认新密码" : "비밀번호 확인"}
                  style={{ width: "100%", padding: "9px 12px", borderRadius: 8, boxSizing: "border-box", fontSize: 13,
                    border: editPwConfirm && editPw !== editPwConfirm ? "1.5px solid #f87171" : "1px solid #d1d5db" }} />
                {editPwConfirm && (
                  <div style={{ fontSize: 11, marginTop: 4, color: editPw === editPwConfirm ? "#10b981" : "#dc2626" }}>
                    {editPw === editPwConfirm
                      ? `✓ ${lang === "zh" ? "密码一致" : "비밀번호 일치"}`
                      : lang === "zh" ? "密码不一致" : "비밀번호가 일치하지 않습니다"}
                  </div>
                )}
              </div>
            )}

            {editError && <div style={{ fontSize: 12, color: "#dc2626", marginBottom: 8 }}>{editError}</div>}

            <div style={{ display: "flex", gap: 10, marginTop: 20, justifyContent: "flex-end" }}>
              <button onClick={() => setEditTarget(null)}
                style={{ padding: "9px 20px", borderRadius: 8, border: "1px solid #d1d5db", background: "#f9fafb", cursor: "pointer" }}>{t.cancel}</button>
              <button onClick={saveEdit}
                style={{ padding: "9px 20px", borderRadius: 8, border: "none", background: "#1a56db", color: "#fff", fontWeight: 600, cursor: "pointer" }}>{t.save}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// 主应用
// ============================================================
export default function App() {
  const [lang, setLang] = useState("zh");
  const [page, setPage] = useState("dashboard");
  const [users, setUsers] = useState(MOCK_USERS);
  const [currentUser, setCurrentUser] = useState(null);
  const [quotes, setQuotes] = useState(MOCK_QUOTES);
  const [orders, setOrders] = useState(MOCK_ORDERS);
  const [quoteModal, setQuoteModal] = useState(null);
  const [orderModal, setOrderModal] = useState(null);
  const [exportModal, setExportModal] = useState(false);
  const [globalRates, setGlobalRates] = useState({ CNY: 190.5, USD: 1382.0 });
  const [filterStatus, setFilterStatus] = useState("all");
  const [filterCustomer, setFilterCustomer] = useState("all");
  const [filterStaff, setFilterStaff] = useState("all");
  const [filterCategory, setFilterCategory] = useState("all");
  const [searchTerm, setSearchTerm] = useState("");
  const t = T[lang];
  const perm = currentUser ? PERMISSIONS[currentUser.role] : null;

  const customers = useMemo(() => [...new Set([...quotes, ...orders].map(c => c.customer))], [quotes, orders]);
  const staff = useMemo(() => [...new Set([...quotes, ...orders].flatMap(c => [c.salesRep, c.purchaseRep]).filter(Boolean))], [quotes, orders]);

  const filterFn = c =>
    (filterStatus === "all" || c.status === filterStatus) &&
    (filterCustomer === "all" || c.customer === filterCustomer) &&
    (filterStaff === "all" || c.salesRep === filterStaff || c.purchaseRep === filterStaff) &&
    (filterCategory === "all" || c.category === filterCategory || !c.category) &&
    (!searchTerm.trim() || JSON.stringify(c).toLowerCase().includes(searchTerm.trim().toLowerCase()));

  const filteredQuotes = quotes.filter(filterFn);
  const filteredOrders = orders.filter(filterFn);

  function saveQuote(form) {
    setQuotes(qs => qs.some(q => q.id === form.id) ? qs.map(q => q.id === form.id ? form : q) : [...qs, form]);
    setQuoteModal(null);
  }
  function saveOrder(form) {
    setOrders(os => os.some(o => o.id === form.id) ? os.map(o => o.id === form.id ? form : o) : [...os, form]);
    setOrderModal(null);
  }
  function deleteQuote(id) { if (confirm("삭제하시겠습니까? / 确定删除？")) setQuotes(qs => qs.filter(q => q.id !== id)); }
  function deleteOrder(id) { if (confirm("삭제하시겠습니까? / 确定删除？")) setOrders(os => os.filter(o => o.id !== id)); }
  function quickStatus(id, status) {
    setQuotes(qs => qs.map(q => q.id === id ? { ...q, status, statusDate: new Date().toISOString().slice(0,10) } : q));
  }

  // 转下单：每个产品取「选定」的比价行，各生成一张下单案件
  function convertToOrder(q) {
    const newOrders = q.items.map((it, i) => {
      const sel = it.sources.find(s => s.selected);
      return {
        id: `O${new Date().getFullYear()}-${String(Math.floor(Math.random()*900)+100 + i)}`,
        linkedQuote: q.id, customer: q.customer,
        product: it.productName, spec: it.detailQty || it.colorSize, qty: sel?.qty || "",
        purchaseCurrency: sel?.currency || "CNY",
        purchaseAmount: (parseFloat(sel?.unitPrice) || 0) * (parseFloat(sel?.qty) || 1),
        inlandFee: (parseFloat(sel?.inlandFee) || 0) + (parseFloat(sel?.packingFee) || 0),
        lockedRate: sel?.lockedRate || LIVE_RATES.CNY,
        salePrice: it.salePrice, supplier: sel?.factory || "",
        salesRep: q.salesRep, purchaseRep: q.purchaseRep, category: q.category,
        status: "os1", paymentStatus: "미결제",
        deliveryDate: q.estimatedDelivery,
        factoryDelivery: "", estShipDate: "", warehouseDate: "", loadingDate: "",
        shipInfo: "", arrivalDate: "", customsClear: "",
        logistics: sel ? `${t.packing}: ${sel.packing || "—"} / CBM: ${sel.cbm || "—"} / ${sel.weightKg || "—"}kg` : "",
        customerNote: it.customerReq
      };
    });
    setOrders(os => [...os, ...newOrders]);
    setPage("order");
  }

  // 未登录 → 登录页
  if (!currentUser) {
    return <LoginScreen users={users} lang={lang} setLang={setLang} onLogin={u => { setCurrentUser(u); setPage("dashboard"); }} />;
  }

  const navItems = [
    { key: "dashboard", label: t.dashboard, icon: "🏠" },
    { key: "quote", label: t.quoteSystem, icon: "📋" },
    { key: "order", label: t.orderSystem, icon: "📦" },
    ...(currentUser.role === "admin" ? [{ key: "staff", label: t.staffMgmt, icon: "👥" }] : []),
  ];
  const STATUSES_FOR_PAGE = page === "quote" ? QUOTE_STATUSES : ORDER_STATUSES;

  return (
    <div style={{ display: "flex", minHeight: "100vh", background: "#f8fafc", fontFamily: "'Noto Sans KR', 'PingFang SC', sans-serif" }}>
      {/* Sidebar */}
      <div style={{ width: 230, background: "#0f172a", display: "flex", flexDirection: "column", padding: "24px 0", flexShrink: 0 }}>
        <div style={{ padding: "0 20px 20px", borderBottom: "1px solid rgba(255,255,255,0.08)" }}>
          <div style={{ fontWeight: 700, fontSize: 13, color: "#fff", lineHeight: 1.4 }}>{t.appTitle}</div>
          <div style={{ fontSize: 11, color: "#64748b", marginTop: 4 }}>Trade Order System v3</div>
        </div>
        <nav style={{ padding: "14px 0", flex: 1 }}>
          {navItems.map(item => (
            <button key={item.key} onClick={() => { setPage(item.key); setFilterStatus("all"); setFilterCustomer("all"); setFilterStaff("all"); setFilterCategory("all"); setSearchTerm(""); }}
              style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "10px 20px", border: "none", background: page === item.key ? "rgba(59,130,246,0.15)" : "transparent", color: page === item.key ? "#60a5fa" : "#94a3b8", fontSize: 13, cursor: "pointer", borderLeft: page === item.key ? "3px solid #3b82f6" : "3px solid transparent", textAlign: "left" }}>
              <span>{item.icon}</span>{item.label}
            </button>
          ))}
        </nav>
        <div style={{ padding: "12px 20px", borderTop: "1px solid rgba(255,255,255,0.08)" }}>
          <div style={{ fontSize: 10, color: "#475569", marginBottom: 8 }}>
            📡 {lang === "zh" ? "参考汇率 (管理员可修改)" : "참고 환율 (관리자 수정 가능)"} → KRW
          </div>
          {["CNY","USD"].map(c => (
            <div key={c} style={{ display:"flex", alignItems:"center", gap:6, marginBottom:6 }}>
              <span style={{ fontSize:11, color:"#64748b", width:28 }}>{c}</span>
              <input type="number" step="0.1"
                value={globalRates[c]}
                onChange={e => {
                  const v = parseFloat(e.target.value) || 0;
                  setGlobalRates(r => ({...r, [c]: v}));
                  LIVE_RATES[c] = v;
                }}
                style={{ flex:1, padding:"3px 6px", borderRadius:5, border:"1px solid rgba(255,255,255,0.2)", background:"#1e293b", color:"#e2e8f0", fontSize:12, width:0 }} />
            </div>
          ))}
        </div>
        <div style={{ padding: "12px 20px", borderTop: "1px solid rgba(255,255,255,0.08)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 34, height: 34, borderRadius: "50%", background: "#1e3a8a", color: "#93c5fd", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, flexShrink: 0 }}>
              {currentUser.name.slice(0, 1)}
            </div>
            <div style={{ minWidth: 0 }}>
              <div style={{ fontSize: 13, color: "#e2e8f0", fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{currentUser.name}</div>
              <div style={{ fontSize: 10, color: "#64748b" }}>{PERMISSIONS[currentUser.role].label[lang]} · {currentUser.id}</div>
            </div>
          </div>
          <button onClick={() => setCurrentUser(null)}
            style={{ width: "100%", marginTop: 10, padding: "7px 12px", borderRadius: 8, border: "1px solid rgba(248,113,113,0.3)", background: "rgba(248,113,113,0.08)", color: "#fca5a5", fontSize: 12, cursor: "pointer" }}>
            ⎋ {t.logout}
          </button>
          <button onClick={() => setLang(l => l === "zh" ? "ko" : "zh")}
            style={{ width: "100%", marginTop: 8, padding: "7px 12px", borderRadius: 8, border: "1px solid rgba(255,255,255,0.15)", background: "rgba(255,255,255,0.06)", color: "#94a3b8", fontSize: 12, cursor: "pointer" }}>
            🌐 {t.lang}
          </button>
        </div>
      </div>

      {/* Main */}
      <div style={{ flex: 1, padding: 28, overflowY: "auto", minWidth: 0 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
          <div>
            <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, color: "#0f172a" }}>
              {page === "dashboard" ? t.dashboard : page === "quote" ? t.quoteSystem : page === "order" ? t.orderSystem : t.staffMgmt}
            </h1>
            <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 3 }}>
              👤 {currentUser.name} · <b style={{ color: "#1a56db" }}>{perm.label[lang]}</b>
            </div>
          </div>
          <div style={{ display: "flex", gap: 10 }}>
            {(page === "quote" || page === "order") && (
              <button onClick={() => setExportModal(true)}
                style={{ padding: "9px 16px", borderRadius: 9, border: "1px solid #d1d5db", background: "#fff", fontSize: 13, cursor: "pointer", color: "#374151" }}>
                📥 {t.exportExcel}
              </button>
            )}
            {page !== "dashboard" && page !== "staff" && (
              <button onClick={() => page === "quote" ? setQuoteModal("new") : setOrderModal("new")}
                style={{ padding: "9px 18px", borderRadius: 9, border: "none", background: "#1a56db", color: "#fff", fontWeight: 600, fontSize: 13, cursor: "pointer" }}>
                + {page === "quote" ? t.newQuote : t.newOrder}
              </button>
            )}
          </div>
        </div>

        {page === "quote" && (
          <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
            {[["all", t.allCategories, "📋"], ["daily", t.catDaily, "🧺"], ["machine", t.catMachine, "⚙️"]].map(([k, label, icon]) => {
              const count = k === "all" ? quotes.length : quotes.filter(q => q.category === k).length;
              const active = filterCategory === k;
              return (
                <button key={k} onClick={() => setFilterCategory(k)}
                  style={{ padding: "8px 18px", borderRadius: 9, border: active ? "2px solid #1a56db" : "1px solid #d1d5db", background: active ? "#eff6ff" : "#fff", color: active ? "#1a56db" : "#6b7280", fontSize: 13, fontWeight: active ? 700 : 400, cursor: "pointer" }}>
                  {icon} {label} <span style={{ fontSize: 11, opacity: 0.7 }}>({count})</span>
                </button>
              );
            })}
          </div>
        )}

        {page !== "dashboard" && page !== "staff" && (
          <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap" }}>
            <input value={searchTerm} onChange={e => setSearchTerm(e.target.value)} placeholder={t.searchPlaceholder}
              style={{ flex: "1 1 260px", minWidth: 220, padding: "7px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 12, boxSizing: "border-box" }} />
            <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 12, background: "#fff" }}>
              <option value="all">{t.allStatus}</option>
              {STATUSES_FOR_PAGE.map(s => <option key={s} value={s}>{t[s]}</option>)}
            </select>
            <select value={filterCustomer} onChange={e => setFilterCustomer(e.target.value)} style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 12, background: "#fff" }}>
              <option value="all">{t.allCustomers}</option>
              {customers.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
            <select value={filterStaff} onChange={e => setFilterStaff(e.target.value)} style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid #d1d5db", fontSize: 12, background: "#fff" }}>
              <option value="all">{t.allStaff}</option>
              {staff.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
        )}

        {page === "dashboard" && <Dashboard quotes={quotes} orders={orders} lang={lang} perm={perm} onNav={setPage} />}
        {page === "staff" && currentUser.role === "admin" && <StaffPage users={users} setUsers={setUsers} currentUser={currentUser} lang={lang} perm={perm} />}
        {page === "quote" && (
          <div style={{ maxWidth: 760 }}>
            {filteredQuotes.map(q => (
              <QuoteCard key={q.id} q={q} lang={lang} perm={perm} user={currentUser}
                onDetail={r => setQuoteModal(r)} onConvert={convertToOrder} onDelete={deleteQuote} onQuickStatus={quickStatus} />
            ))}
          </div>
        )}
        {page === "order" && (
          <div style={{ columns: "2", columnGap: 16 }}>
            {filteredOrders.map(o => (
              <div key={o.id} style={{ breakInside: "avoid" }}>
                <OrderCard o={o} lang={lang} perm={perm} user={currentUser}
                  onEdit={r => setOrderModal(r)} onDelete={deleteOrder}
                  onAdvance={(id, st) => setOrders(os => os.map(x => x.id === id ? { ...x, status: st } : x))} />
              </div>
            ))}
          </div>
        )}
      </div>

      {exportModal && (
        <ExportModal
          quotes={quotes} orders={orders}
          filteredQuotes={filteredQuotes} filteredOrders={filteredOrders}
          perm={perm} lang={lang} onClose={() => setExportModal(false)} />
      )}
      {quoteModal && <QuoteDetailModal initial={quoteModal === "new" ? null : quoteModal} lang={lang} perm={perm} user={currentUser} onSave={saveQuote} onClose={() => setQuoteModal(null)} />}
      {orderModal && <OrderModal initial={orderModal === "new" ? null : orderModal} lang={lang} perm={perm} user={currentUser} onSave={saveOrder} onClose={() => setOrderModal(null)} />}
    </div>
  );
}


// Mount
const container = document.getElementById("root");
if (container && !container._reactRoot) {
  container._reactRoot = true;
  createRoot(container).render(<App />);
}
