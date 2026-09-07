import { Fragment, useEffect, useState } from "react";
import { Plus, Trash2, FileSpreadsheet, RotateCcw, Eye, X, Mail } from "lucide-react";
import ExcelJS from "exceljs";

let uidCounter = 1;
const uid = () => uidCounter++;
const advanceUidPast = (id) => {
  if (typeof id === "number" && id >= uidCounter) uidCounter = id + 1;
};

const HEADER_STORAGE_KEY = "quote-app-header";
const GROUPS_STORAGE_KEY = "quote-app-groups";
const EMAIL_STORAGE_KEY = "quote-app-recipient-email";

const initialHeader = {
  productName: "",
  regNo: "",
  styleNo: "",
  company: "",
  ceoName: "",
  orderQty: "",
  address: "",
  color: "",
  bizType: "",
  category: "",
};

const loadStoredHeader = () => {
  try {
    const raw = localStorage.getItem(HEADER_STORAGE_KEY);
    if (!raw) return initialHeader;
    return { ...initialHeader, ...JSON.parse(raw) };
  } catch {
    return initialHeader;
  }
};

const seedItem = (name, unit, price, qty, amortize) => ({
  id: uid(),
  name,
  unit,
  price: String(price),
  qty: String(qty),
  amortize: !!amortize,
});

const initialGroups = [
  { id: uid(), major: "원자재", sub: "외피", items: [] },
  { id: uid(), major: "원자재", sub: "내피", items: [] },
  { id: uid(), major: "부자재", sub: "지부재", items: [] },
  { id: uid(), major: "부자재", sub: "보강재", items: [] },
  { id: uid(), major: "부자재", sub: "기타", items: [] },
  { id: uid(), major: "추가비용", sub: "가공비", items: [] },
];

const loadStoredGroups = () => {
  try {
    const raw = localStorage.getItem(GROUPS_STORAGE_KEY);
    if (!raw) return initialGroups;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return initialGroups;
    parsed.forEach((g) => {
      advanceUidPast(g.id);
      (g.items || []).forEach((it) => advanceUidPast(it.id));
    });
    return parsed;
  } catch {
    return initialGroups;
  }
};

const num = (v) => {
  const n = parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
};

const won = (v) =>
  Math.round(v).toLocaleString("ko-KR", { maximumFractionDigits: 0 });

function itemAmount(item, orderQty) {
  const amt = num(item.price) * num(item.qty);
  return item.amortize ? amt / (num(orderQty) || 1) : amt;
}

// Flattens groups into item rows, marking where the major/sub columns
// should start a merged span. Shared by the Excel export and the preview
// so both always show exactly the same layout.
function buildQuoteRows(groups) {
  const rows = [];
  const groupMeta = [];
  groups.forEach((g) => {
    const items = g.items.length ? g.items : [null];
    const startIndex = rows.length;
    items.forEach((item) => {
      rows.push({ major: g.major, sub: g.sub, item });
    });
    rows[startIndex].groupStart = true;
    rows[startIndex].groupSpan = items.length;
    groupMeta.push({ major: g.major, startIndex, rowCount: items.length });
  });

  let i = 0;
  while (i < groupMeta.length) {
    let j = i;
    let total = groupMeta[i].rowCount;
    while (j + 1 < groupMeta.length && groupMeta[j + 1].major === groupMeta[i].major) {
      j++;
      total += groupMeta[j].rowCount;
    }
    rows[groupMeta[i].startIndex].majorStart = true;
    rows[groupMeta[i].startIndex].majorSpan = total;
    i = j + 1;
  }
  return rows;
}

export default function App() {
  const [header, setHeader] = useState(loadStoredHeader);
  const [groups, setGroups] = useState(loadStoredGroups);
  const [freight, setFreight] = useState("1000");
  const [marginRate, setMarginRate] = useState("15");
  const [showPreview, setShowPreview] = useState(false);
  const [emailCopied, setEmailCopied] = useState(false);
  const [recipientEmail, setRecipientEmail] = useState(() => {
    try {
      return localStorage.getItem(EMAIL_STORAGE_KEY) || "";
    } catch {
      return "";
    }
  });

  useEffect(() => {
    if (!emailCopied) return;
    const t = setTimeout(() => setEmailCopied(false), 4000);
    return () => clearTimeout(t);
  }, [emailCopied]);

  useEffect(() => {
    try {
      localStorage.setItem(HEADER_STORAGE_KEY, JSON.stringify(header));
    } catch {
      // ignore storage failures (e.g. private browsing)
    }
  }, [header]);

  useEffect(() => {
    try {
      localStorage.setItem(GROUPS_STORAGE_KEY, JSON.stringify(groups));
    } catch {
      // ignore storage failures (e.g. private browsing)
    }
  }, [groups]);

  useEffect(() => {
    try {
      localStorage.setItem(EMAIL_STORAGE_KEY, recipientEmail);
    } catch {
      // ignore storage failures (e.g. private browsing)
    }
  }, [recipientEmail]);

  const setHeaderField = (key) => (e) =>
    setHeader((h) => ({ ...h, [key]: e.target.value }));

  const updateItem = (groupId, itemId, key, value) => {
    setGroups((gs) =>
      gs.map((g) =>
        g.id !== groupId
          ? g
          : {
              ...g,
              items: g.items.map((it) =>
                it.id !== itemId ? it : { ...it, [key]: value }
              ),
            }
      )
    );
  };

  const addItem = (groupId) => {
    setGroups((gs) =>
      gs.map((g) =>
        g.id !== groupId
          ? g
          : { ...g, items: [...g.items, seedItem("", "족", "", "")] }
      )
    );
  };

  const removeItem = (groupId, itemId) => {
    setGroups((gs) =>
      gs.map((g) =>
        g.id !== groupId
          ? g
          : { ...g, items: g.items.filter((it) => it.id !== itemId) }
      )
    );
  };

  const resetAll = () => {
    setHeader(initialHeader);
    setGroups(initialGroups);
    setFreight("1000");
    setMarginRate("15");
  };

  const groupSubtotal = (g) =>
    g.items.reduce((sum, it) => sum + itemAmount(it, header.orderQty), 0);

  const productionCost = groups.reduce((sum, g) => sum + groupSubtotal(g), 0);
  const freightNum = num(freight);
  const marginAmount = (productionCost + freightNum) * (num(marginRate) / 100);
  const supplyAmount = productionCost + freightNum + marginAmount;

  const exportToExcel = async () => {
    const FONT = "맑은 고딕";
    const GRAY = "FFD9D9D9";
    const LGRAY = "FFF2F2F2";
    const YELLOW = "FFFFFF99";
    const BLUE = "FF0000FF";
    const thin = { style: "thin" };
    const allBorder = { top: thin, bottom: thin, left: thin, right: thin };

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("견적서", { views: [{ state: "frozen", ySplit: 8 }] });
    ws.columns = [
      { width: 10 },
      { width: 12 },
      { width: 16 },
      { width: 10 },
      { width: 12 },
      { width: 10 },
      { width: 13 },
    ];

    const setCell = (addr, value, opts = {}) => {
      const cell = ws.getCell(addr);
      cell.value = value;
      cell.font = {
        name: FONT,
        size: opts.size || 11,
        bold: !!opts.bold,
        color: opts.color ? { argb: opts.color } : undefined,
      };
      cell.alignment = {
        horizontal: opts.align || "center",
        vertical: "middle",
        wrapText: true,
      };
      cell.border = allBorder;
      if (opts.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
      if (opts.numFmt) cell.numFmt = opts.numFmt;
      return cell;
    };

    const label = (addr, text, opts = {}) => setCell(addr, text, { bold: true, fill: GRAY, ...opts });
    const input = (addr, value, opts = {}) =>
      setCell(addr, value, { fill: YELLOW, color: opts.numeric ? BLUE : undefined, ...opts });

    // Title
    ws.mergeCells("A1:G1");
    setCell("A1", "견 적 서", { size: 20, bold: true });
    ws.getRow(1).height = 36;

    // Header info block
    ws.mergeCells("B3:C3");
    ws.mergeCells("E3:G3");
    label("A3", "품명");
    input("B3", header.productName);
    label("D3", "등록번호");
    input("E3", header.regNo);

    ws.mergeCells("B4:C4");
    label("A4", "STYLE No");
    input("B4", header.styleNo);
    label("D4", "상호");
    input("E4", header.company);
    label("F4", "성명");
    input("G4", header.ceoName);

    ws.mergeCells("B5:C5");
    ws.mergeCells("E5:G5");
    label("A5", "발주량");
    input("B5", num(header.orderQty), { numeric: true, numFmt: "#,##0" });
    label("D5", "사업장주소");
    input("E5", header.address, { align: "left" });

    ws.mergeCells("B6:C6");
    label("A6", "COLOR");
    input("B6", header.color);
    label("D6", "업태");
    input("E6", header.bizType);
    label("F6", "종목");
    input("G6", header.category);

    ws.getRow(3).height = 20;
    ws.getRow(4).height = 20;
    ws.getRow(5).height = 26;
    ws.getRow(6).height = 20;

    // Column headers
    ws.mergeCells("A8:C8");
    label("A8", "구분");
    label("D8", "소재");
    label("E8", "단가");
    label("F8", "소요량");
    label("G8", "금액");

    // Item rows, grouped by major/sub like the reference sheet
    let r = 9;
    const quoteRows = buildQuoteRows(groups);
    quoteRows.forEach((row) => {
      const it = row.item;
      if (it) {
        setCell(`C${r}`, it.name, { align: "left" });
        setCell(`D${r}`, it.unit);
        input(`E${r}`, num(it.price), { numeric: true, align: "right", numFmt: "#,##0" });
        input(`F${r}`, num(it.qty), { numeric: true, align: "right", numFmt: "#,##0.00" });
        const formula = it.amortize ? `E${r}*F${r}/$B$5` : `E${r}*F${r}`;
        setCell(`G${r}`, { formula }, { align: "right", numFmt: "#,##0" });
      } else {
        setCell(`C${r}`, "");
        setCell(`D${r}`, "");
        setCell(`E${r}`, "");
        setCell(`F${r}`, "");
        setCell(`G${r}`, "");
      }
      if (row.groupStart) {
        const end = r + row.groupSpan - 1;
        if (r !== end) ws.mergeCells(`B${r}:B${end}`);
        label(`B${r}`, row.sub);
      }
      if (row.majorStart) {
        const end = r + row.majorSpan - 1;
        if (r !== end) ws.mergeCells(`A${r}:A${end}`);
        label(`A${r}`, row.major);
      }
      r++;
    });
    const lastItemRow = r - 1;

    // Totals
    const pcRow = r;
    ws.mergeCells(`A${pcRow}:F${pcRow}`);
    label(`A${pcRow}`, "생산원가", { fill: LGRAY });
    setCell(`G${pcRow}`, { formula: `SUM(G9:G${lastItemRow})` }, { bold: true, fill: LGRAY, align: "right", numFmt: "#,##0" });
    r++;

    const frRow = r;
    ws.mergeCells(`A${frRow}:F${frRow}`);
    label(`A${frRow}`, "운임비", { fill: LGRAY });
    input(`G${frRow}`, freightNum, { numeric: true, align: "right", numFmt: "#,##0" });
    r++;

    const mgRow = r;
    ws.mergeCells(`A${mgRow}:E${mgRow}`);
    label(`A${mgRow}`, "업체마진", { fill: LGRAY });
    input(`F${mgRow}`, num(marginRate) / 100, { numeric: true, numFmt: "0%" });
    setCell(`G${mgRow}`, { formula: `(G${pcRow}+G${frRow})*F${mgRow}` }, { bold: true, fill: LGRAY, align: "right", numFmt: "#,##0" });
    r += 2; // spacer row

    // DESIGN / 기타사항 / 공급가액 block
    const designRow = r;
    ws.mergeCells(`A${designRow}:B${designRow + 1}`);
    label(`A${designRow}`, "DESIGN");
    ws.mergeCells(`C${designRow}:D${designRow}`);
    input(`C${designRow}`, "");
    ws.mergeCells(`C${designRow + 1}:D${designRow + 1}`);
    input(`C${designRow + 1}`, "");
    ws.mergeCells(`E${designRow}:G${designRow}`);
    label(`E${designRow}`, "기타사항");
    label(`E${designRow + 1}`, "공급가액");
    ws.mergeCells(`F${designRow + 1}:G${designRow + 1}`);
    setCell(
      `F${designRow + 1}`,
      { formula: `G${pcRow}+G${frRow}+G${mgRow}` },
      { bold: true, fill: LGRAY, align: "right", numFmt: "#,##0" }
    );

    const noteRow = designRow + 2;
    ws.mergeCells(`A${noteRow}:G${noteRow}`);
    setCell(
      `A${noteRow}`,
      "※ 노란색 칸: 단가/소요량/운임비/마진율 입력란 (수정 가능) · 금액·생산원가·업체마진·공급가액은 자동 계산됩니다.",
      { align: "left", size: 9 }
    );
    ws.getRow(noteRow).height = 26;

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = header.productName ? `견적서_${header.productName}.xlsx` : "견적서.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const buildEmailContent = () => {
    const subject = `견적서${header.productName ? ` - ${header.productName}` : ""}${
      header.company ? ` (${header.company})` : ""
    }`;
    const body = [
      header.company ? `${header.company} 담당자님께,` : "안녕하세요,",
      "",
      "견적서를 보내드립니다. 첨부된 엑셀 파일을 확인해 주세요.",
      "",
      `- 품명: ${header.productName || "-"}`,
      `- STYLE No: ${header.styleNo || "-"}`,
      `- 발주량: ${header.orderQty || "-"}`,
      `- 공급가액: ${won(supplyAmount)}원`,
      "",
      "감사합니다.",
    ].join("\n");
    return { subject, body };
  };

  // Copying to the clipboard must happen while the page still has focus.
  // Any subsequent mailto click or window.open hands focus to the OS/a new
  // tab, and the Clipboard API silently fails once focus is gone — so this
  // always runs first, before any navigation.
  const copyEmailToClipboard = (subject, body) => {
    if (!navigator.clipboard?.writeText) return;
    navigator.clipboard
      .writeText(`제목: ${subject}\n\n${body}`)
      .then(() => setEmailCopied(true))
      .catch(() => {});
  };

  const sendByEmail = () => {
    // Everything here must run synchronously, as a direct result of the
    // click — an await or alert first makes mobile browsers stop treating
    // the mailto navigation as user-initiated and silently drop it (no app
    // chooser, nothing happens).
    const { subject, body } = buildEmailContent();
    copyEmailToClipboard(subject, body);

    const to = recipientEmail.trim();
    const mailtoUrl = `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

    // Clicking a real <a href="mailto:..."> is more reliably picked up by
    // mobile browsers than assigning window.location.href directly. Not
    // every mail app (e.g. 네이버메일) registers itself as a mailto handler
    // though, so the OS chooser may only ever offer Gmail or the phone's
    // built-in Mail app — the clipboard copy above is the fallback for that.
    const link = document.createElement("a");
    link.href = mailtoUrl;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Kick off the download after — it doesn't need the user-gesture context.
    exportToExcel();
  };

  const sendByNaverMail = () => {
    // Naver Mail doesn't offer a documented way to prefill recipient/subject/
    // body via URL, so instead: copy the text, open the (always-correct)
    // Naver Mail inbox in a new tab, and let the user paste it themselves.
    const { subject, body } = buildEmailContent();
    copyEmailToClipboard(subject, body);
    window.open("https://mail.naver.com/", "_blank", "noopener");
    exportToExcel();
  };

  return (
    <div className="min-h-screen bg-stone-100 text-stone-800 text-base pb-28 print:bg-white print:pb-0">
      <div className="max-w-4xl mx-auto px-4 py-8 print:p-0 print:max-w-none">
        {/* Title */}
        <div className="border-b-4 border-stone-800 pb-3 mb-6">
          <h1 className="font-serif text-4xl font-bold tracking-wide text-stone-900">
            견 적 서
          </h1>
          <p className="text-sm text-stone-500 mt-1 print:hidden">
            노란 칸에 숫자를 입력하면 금액이 자동으로 계산됩니다.
          </p>
        </div>

        {/* Header info: one field per line so every label/input lines up in a column */}
        <div className="bg-white border border-stone-300 rounded-lg mb-6 overflow-hidden divide-y divide-stone-200 print:border-stone-800">
          <InfoField label="품명" value={header.productName} onChange={setHeaderField("productName")} />
          <InfoField label="등록번호" value={header.regNo} onChange={setHeaderField("regNo")} />
          <InfoField label="STYLE No" value={header.styleNo} onChange={setHeaderField("styleNo")} />
          <InfoField label="상호" value={header.company} onChange={setHeaderField("company")} />
          <InfoField label="성명" value={header.ceoName} onChange={setHeaderField("ceoName")} />
          <InfoField label="발주량" value={header.orderQty} onChange={setHeaderField("orderQty")} accent />
          <InfoField label="사업장주소" value={header.address} onChange={setHeaderField("address")} />
          <InfoField label="COLOR" value={header.color} onChange={setHeaderField("color")} />
          <InfoField label="업태" value={header.bizType} onChange={setHeaderField("bizType")} />
          <InfoField label="종목" value={header.category} onChange={setHeaderField("category")} />
        </div>

        {/* Item groups: a real table per group, columns matching the Excel sheet */}
        <div className="space-y-6">
          {groups.map((g) => (
            <div key={g.id} className="bg-white border border-stone-300 rounded-lg overflow-hidden print:border-stone-800">
              <div className="flex items-center justify-between bg-stone-800 text-stone-50 px-4 py-3">
                <div className="font-serif text-lg tracking-wide">
                  {g.major} <span className="text-stone-300 mx-1">/</span> {g.sub}
                </div>
                <div className="text-lg font-semibold tabular-nums">
                  {won(groupSubtotal(g))}원
                </div>
              </div>

              {/* Desktop/tablet: a real table matching the Excel columns */}
              <div className="hidden sm:block overflow-x-auto">
                <table className="w-full border-collapse text-base">
                  <thead>
                    <tr className="bg-stone-100 text-stone-600">
                      <th className="border border-stone-300 py-2.5 px-2 text-left font-medium min-w-[140px]">품목</th>
                      <th className="border border-stone-300 py-2.5 px-2 font-medium min-w-[90px]">소재</th>
                      <th className="border border-stone-300 py-2.5 px-2 font-medium min-w-[110px]">단가</th>
                      <th className="border border-stone-300 py-2.5 px-2 font-medium min-w-[90px]">소요량</th>
                      <th className="border border-stone-300 py-2.5 px-2 font-medium min-w-[110px]">금액</th>
                      <th className="border border-stone-300 py-2.5 px-2 w-12 print:hidden"></th>
                    </tr>
                  </thead>
                  <tbody>
                    {g.items.length === 0 && (
                      <tr>
                        <td colSpan={6} className="border border-stone-300 py-4 text-center text-stone-400">
                          아래 버튼으로 항목을 추가하세요
                        </td>
                      </tr>
                    )}
                    {g.items.map((it) => (
                      <Fragment key={it.id}>
                        <tr>
                          <td className="border border-stone-300 p-1">
                            <input
                              className="w-full bg-transparent outline-none px-2 py-2 text-base"
                              value={it.name}
                              placeholder="품목명"
                              onChange={(e) => updateItem(g.id, it.id, "name", e.target.value)}
                            />
                          </td>
                          <td className="border border-stone-300 p-1">
                            <input
                              className="w-full bg-transparent outline-none px-2 py-2 text-base text-center"
                              value={it.unit}
                              placeholder="단위"
                              onChange={(e) => updateItem(g.id, it.id, "unit", e.target.value)}
                            />
                          </td>
                          <td className="border border-stone-300 p-1">
                            <input
                              className="w-full bg-amber-50 outline-none px-2 py-2 text-base text-right tabular-nums"
                              inputMode="decimal"
                              value={it.price}
                              onChange={(e) => updateItem(g.id, it.id, "price", e.target.value)}
                            />
                          </td>
                          <td className="border border-stone-300 p-1">
                            <input
                              className="w-full bg-amber-50 outline-none px-2 py-2 text-base text-right tabular-nums"
                              inputMode="decimal"
                              value={it.qty}
                              onChange={(e) => updateItem(g.id, it.id, "qty", e.target.value)}
                            />
                          </td>
                          <td className="border border-stone-300 px-2 py-2 text-right font-medium tabular-nums">
                            {won(itemAmount(it, header.orderQty))}
                          </td>
                          <td className="border border-stone-300 text-center print:hidden">
                            <button
                              onClick={() => removeItem(g.id, it.id)}
                              className="text-stone-400 hover:text-red-500 p-2"
                              aria-label="행 삭제"
                            >
                              <Trash2 size={20} />
                            </button>
                          </td>
                        </tr>
                        <tr className="print:hidden">
                          <td colSpan={6} className="border border-stone-300 border-t-0 px-3 py-1.5 bg-stone-50">
                            <label className="flex items-center gap-2 text-sm text-stone-600">
                              <input
                                type="checkbox"
                                checked={it.amortize}
                                onChange={(e) => updateItem(g.id, it.id, "amortize", e.target.checked)}
                                className="w-4 h-4 accent-amber-700"
                              />
                              발주량({header.orderQty || 0})으로 나눠서 계산 (금형비 등 1회성 비용일 때 체크)
                            </label>
                          </td>
                        </tr>
                      </Fragment>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Mobile: one field per line, no side-scrolling */}
              <div className="sm:hidden p-3 space-y-3">
                {g.items.length === 0 && (
                  <div className="text-center text-stone-400 py-3">아래 버튼으로 항목을 추가하세요</div>
                )}
                {g.items.map((it) => (
                  <MobileItemCard
                    key={it.id}
                    item={it}
                    orderQty={header.orderQty}
                    onChange={(key, value) => updateItem(g.id, it.id, key, value)}
                    onRemove={() => removeItem(g.id, it.id)}
                  />
                ))}
              </div>

              <button
                onClick={() => addItem(g.id)}
                className="w-full flex items-center justify-center gap-2 text-base text-stone-600 hover:text-stone-900 hover:bg-stone-50 py-3 border-t border-stone-200 print:hidden"
              >
                <Plus size={18} /> {g.sub}에 항목 추가
              </button>
            </div>
          ))}
        </div>

        {/* Totals */}
        <div className="mt-6 bg-white border border-stone-300 rounded-lg p-4 print:border-stone-800">
          <TotalRow label="생산원가" value={won(productionCost)} />
          <div className="flex items-center justify-between py-3 border-t border-stone-100">
            <span className="text-base text-stone-600">운임비</span>
            <input
              className="w-32 bg-amber-50 border border-amber-200 focus:border-amber-500 rounded-md outline-none px-2 py-2 text-base text-right tabular-nums print:bg-transparent print:border-none"
              inputMode="decimal"
              value={freight}
              onChange={(e) => setFreight(e.target.value)}
            />
          </div>
          <div className="flex items-center justify-between py-3 border-t border-stone-100">
            <span className="text-base text-stone-600">업체마진</span>
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-1">
                <input
                  className="w-20 bg-amber-50 border border-amber-200 focus:border-amber-500 rounded-md outline-none px-2 py-2 text-base text-right tabular-nums print:bg-transparent print:border-none"
                  inputMode="decimal"
                  value={marginRate}
                  onChange={(e) => setMarginRate(e.target.value)}
                />
                <span className="text-base text-stone-500">%</span>
              </div>
              <span className="text-base font-medium tabular-nums w-28 text-right">
                {won(marginAmount)}
              </span>
            </div>
          </div>
          <div className="flex items-center justify-between pt-4 mt-1 border-t-2 border-stone-800">
            <span className="font-serif text-lg text-stone-900">공급가액</span>
            <span className="font-serif text-2xl font-bold tabular-nums text-stone-900">
              {won(supplyAmount)}원
            </span>
          </div>
        </div>

        <div className="mt-5 print:hidden bg-white border border-stone-300 rounded-lg overflow-hidden">
          <InfoField
            label="이메일"
            value={recipientEmail}
            onChange={(e) => setRecipientEmail(e.target.value)}
            type="email"
            placeholder="받는사람 이메일 (예: example@company.com)"
          />
        </div>
        <p className="text-sm text-stone-500 mt-2 print:hidden">
          엑셀 파일이 다운로드되고, 제목·본문은 클립보드에 복사됩니다. 열린 메일 앱(또는 새 탭)에 붙여넣고, 다운로드된 파일을 첨부해서 보내주세요.
        </p>
        {emailCopied && (
          <p className="text-sm text-green-700 bg-green-50 border border-green-200 rounded-md px-3 py-2 mt-2 print:hidden">
            제목과 본문이 복사되었어요. 메일 작성 화면에 붙여넣기 해주세요.
          </p>
        )}

        <div className="flex flex-col sm:flex-row justify-end gap-3 mt-3 print:hidden">
          <button
            onClick={resetAll}
            className="flex items-center justify-center gap-2 text-base text-stone-600 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <RotateCcw size={18} /> 초기값으로
          </button>
          <button
            onClick={() => setShowPreview(true)}
            className="flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <Eye size={18} /> 엑셀 미리보기
          </button>
          <button
            onClick={sendByNaverMail}
            className="flex items-center justify-center gap-2 text-base font-medium text-white px-4 py-3 rounded-lg hover:opacity-90"
            style={{ backgroundColor: "#03C75A" }}
          >
            <Mail size={18} /> 네이버메일로 보내기
          </button>
          <button
            onClick={sendByEmail}
            className="flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <Mail size={18} /> 다른 메일 앱으로
          </button>
          <button
            onClick={exportToExcel}
            className="flex items-center justify-center gap-2 text-base font-medium bg-stone-800 text-white px-6 py-3 rounded-lg hover:bg-stone-700"
          >
            <FileSpreadsheet size={20} /> 엑셀로 저장
          </button>
        </div>
      </div>

      {/* Sticky mobile summary */}
      <div className="fixed bottom-0 inset-x-0 bg-stone-800 text-stone-50 px-4 py-3 flex items-center justify-between sm:hidden print:hidden">
        <span className="text-sm text-stone-300">공급가액</span>
        <span className="font-serif text-xl font-semibold tabular-nums">
          {won(supplyAmount)}원
        </span>
      </div>

      {showPreview && (
        <ExcelPreviewModal
          header={header}
          groups={groups}
          freight={freightNum}
          marginRate={marginRate}
          productionCost={productionCost}
          marginAmount={marginAmount}
          supplyAmount={supplyAmount}
          onClose={() => setShowPreview(false)}
          onDownload={exportToExcel}
        />
      )}
    </div>
  );
}

function InfoField({ label, value, onChange, accent, type = "text", placeholder }) {
  return (
    <div className="flex items-stretch">
      <div className="shrink-0 w-28 sm:w-36 flex items-center justify-center bg-stone-100 text-stone-600 text-sm sm:text-base font-medium px-2 py-3 text-center">
        {label}
      </div>
      <input
        type={type}
        placeholder={placeholder}
        className={
          "flex-1 min-w-0 outline-none px-3 py-3 text-base print:bg-transparent " +
          (accent ? "bg-amber-50 focus:bg-amber-100" : "bg-white focus:bg-stone-50")
        }
        value={value}
        onChange={onChange}
      />
    </div>
  );
}

function MobileItemCard({ item, orderQty, onChange, onRemove }) {
  return (
    <div className="border border-stone-300 rounded-lg overflow-hidden print:hidden">
      <div className="divide-y divide-stone-200">
        <MobileField
          label="품목"
          value={item.name}
          placeholder="품목명"
          onChange={(e) => onChange("name", e.target.value)}
        />
        <MobileField
          label="소재"
          value={item.unit}
          placeholder="단위"
          onChange={(e) => onChange("unit", e.target.value)}
        />
        <MobileField
          label="단가"
          value={item.price}
          onChange={(e) => onChange("price", e.target.value)}
          inputMode="decimal"
          accent
        />
        <MobileField
          label="소요량"
          value={item.qty}
          onChange={(e) => onChange("qty", e.target.value)}
          inputMode="decimal"
          accent
        />
      </div>

      <div className="flex items-center justify-between px-3 py-2.5 bg-stone-50 border-t border-stone-200">
        <span className="text-sm text-stone-500">금액</span>
        <span className="text-lg font-semibold tabular-nums">{won(itemAmount(item, orderQty))}원</span>
      </div>

      <label className="flex items-center gap-2 px-3 py-2.5 text-sm text-stone-600 border-t border-stone-200">
        <input
          type="checkbox"
          checked={item.amortize}
          onChange={(e) => onChange("amortize", e.target.checked)}
          className="w-4 h-4 accent-amber-700 shrink-0"
        />
        발주량({orderQty || 0})으로 나눠서 계산 (1회성 비용일 때 체크)
      </label>

      <button
        onClick={onRemove}
        className="w-full flex items-center justify-center gap-1.5 text-red-500 text-sm py-2.5 border-t border-stone-200 hover:bg-red-50"
      >
        <Trash2 size={16} /> 이 항목 삭제
      </button>
    </div>
  );
}

function MobileField({ label, value, onChange, placeholder, inputMode, accent }) {
  return (
    <div className="flex items-stretch">
      <div className="shrink-0 w-20 flex items-center justify-center bg-stone-100 text-stone-600 text-sm font-medium text-center">
        {label}
      </div>
      <input
        className={
          "flex-1 min-w-0 outline-none px-3 py-2.5 text-base " +
          (accent ? "bg-amber-50 focus:bg-amber-100 text-right tabular-nums" : "bg-white focus:bg-stone-50")
        }
        value={value}
        placeholder={placeholder}
        inputMode={inputMode}
        onChange={onChange}
      />
    </div>
  );
}

function TotalRow({ label, value }) {
  return (
    <div className="flex items-center justify-between py-3">
      <span className="text-base text-stone-600">{label}</span>
      <span className="text-base font-medium tabular-nums">{value}</span>
    </div>
  );
}

function ExcelPreviewModal({
  header,
  groups,
  freight,
  marginRate,
  productionCost,
  marginAmount,
  supplyAmount,
  onClose,
  onDownload,
}) {
  const rows = buildQuoteRows(groups);
  const LABEL = "border border-stone-400 bg-stone-200 font-medium text-center px-2 py-1.5 align-middle";
  const VALUE = "border border-stone-400 bg-yellow-50 px-2 py-1.5 align-middle";
  const PLAIN = "border border-stone-400 bg-white px-2 py-1.5 align-middle";
  const TOTAL_LABEL = "border border-stone-400 bg-stone-100 font-medium text-center px-2 py-1.5 align-middle";
  const TOTAL_VALUE = "border border-stone-400 bg-stone-100 text-right font-semibold px-2 py-1.5 align-middle";

  return (
    <div
      className="fixed inset-0 bg-black/50 flex items-center justify-center p-2 sm:p-6 z-50"
      onClick={onClose}
    >
      <div
        className="bg-white rounded-lg w-full max-w-4xl max-h-full flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3 border-b border-stone-200 shrink-0">
          <h2 className="text-lg font-semibold text-stone-900">엑셀 미리보기</h2>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700 p-1" aria-label="닫기">
            <X size={22} />
          </button>
        </div>

        <div className="overflow-auto p-3 sm:p-4">
          <table className="w-full border-collapse text-sm tabular-nums">
            <tbody>
              <tr>
                <td colSpan={7} className="border border-stone-400 bg-white text-center text-xl font-bold py-3">
                  견 적 서
                </td>
              </tr>
              <tr>
                <td className={LABEL}>품명</td>
                <td colSpan={2} className={VALUE}>{header.productName}</td>
                <td className={LABEL}>등록번호</td>
                <td colSpan={3} className={VALUE}>{header.regNo}</td>
              </tr>
              <tr>
                <td className={LABEL}>STYLE No</td>
                <td colSpan={2} className={VALUE}>{header.styleNo}</td>
                <td className={LABEL}>상호</td>
                <td className={VALUE}>{header.company}</td>
                <td className={LABEL}>성명</td>
                <td className={VALUE}>{header.ceoName}</td>
              </tr>
              <tr>
                <td className={LABEL}>발주량</td>
                <td colSpan={2} className={VALUE + " text-right"}>{header.orderQty}</td>
                <td className={LABEL}>사업장주소</td>
                <td colSpan={3} className={VALUE}>{header.address}</td>
              </tr>
              <tr>
                <td className={LABEL}>COLOR</td>
                <td colSpan={2} className={VALUE}>{header.color}</td>
                <td className={LABEL}>업태</td>
                <td className={VALUE}>{header.bizType}</td>
                <td className={LABEL}>종목</td>
                <td className={VALUE}>{header.category}</td>
              </tr>
              <tr>
                <td colSpan={3} className={LABEL}>구분</td>
                <td className={LABEL}>소재</td>
                <td className={LABEL}>단가</td>
                <td className={LABEL}>소요량</td>
                <td className={LABEL}>금액</td>
              </tr>
              {rows.map((row, idx) => (
                <tr key={idx}>
                  {row.majorStart && (
                    <td rowSpan={row.majorSpan} className={PLAIN + " text-center font-medium"}>
                      {row.major}
                    </td>
                  )}
                  {row.groupStart && (
                    <td rowSpan={row.groupSpan} className={PLAIN + " text-center"}>
                      {row.sub}
                    </td>
                  )}
                  <td className={PLAIN}>{row.item?.name}</td>
                  <td className={PLAIN + " text-center"}>{row.item?.unit}</td>
                  <td className={VALUE + " text-right"}>{row.item ? won(num(row.item.price)) : ""}</td>
                  <td className={VALUE + " text-right"}>{row.item ? row.item.qty : ""}</td>
                  <td className={PLAIN + " text-right"}>
                    {row.item ? won(itemAmount(row.item, header.orderQty)) : ""}
                  </td>
                </tr>
              ))}
              <tr>
                <td colSpan={6} className={TOTAL_LABEL}>생산원가</td>
                <td className={TOTAL_VALUE}>{won(productionCost)}</td>
              </tr>
              <tr>
                <td colSpan={6} className={TOTAL_LABEL}>운임비</td>
                <td className={VALUE + " text-right"}>{won(freight)}</td>
              </tr>
              <tr>
                <td colSpan={5} className={TOTAL_LABEL}>업체마진</td>
                <td className={VALUE + " text-center"}>{marginRate}%</td>
                <td className={TOTAL_VALUE}>{won(marginAmount)}</td>
              </tr>
              <tr>
                <td colSpan={5} rowSpan={2} className={LABEL}>DESIGN</td>
                <td colSpan={2} className="border border-stone-400 bg-yellow-50 px-2 py-3">&nbsp;</td>
              </tr>
              <tr>
                <td className={LABEL}>공급가액</td>
                <td className={TOTAL_VALUE + " text-lg font-bold"}>{won(supplyAmount)}</td>
              </tr>
              <tr>
                <td colSpan={7} className="border border-stone-400 bg-white text-xs text-stone-500 px-2 py-1.5">
                  ※ 노란색 칸: 단가/소요량/운임비/마진율 입력란 (수정 가능) · 금액·생산원가·업체마진·공급가액은 자동 계산됩니다.
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="flex justify-end gap-3 px-4 py-3 border-t border-stone-200 shrink-0">
          <button
            onClick={onClose}
            className="px-4 py-2.5 rounded-lg border border-stone-300 text-stone-600 hover:bg-stone-50"
          >
            닫기
          </button>
          <button
            onClick={onDownload}
            className="flex items-center gap-2 px-5 py-2.5 rounded-lg bg-stone-800 text-white hover:bg-stone-700"
          >
            <FileSpreadsheet size={18} /> 엑셀로 저장
          </button>
        </div>
      </div>
    </div>
  );
}
