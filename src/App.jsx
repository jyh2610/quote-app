import { Fragment, useEffect, useRef, useState } from "react";
import {
  Plus,
  Trash2,
  FileSpreadsheet,
  RotateCcw,
  Eye,
  X,
  Mail,
  Cloud,
  FolderOpen,
  FilePlus2,
  LogOut,
  Upload,
  Share2,
} from "lucide-react";
import ExcelJS from "exceljs";
import { supabase } from "./supabaseClient";

let uidCounter = 1;
const uid = () => uidCounter++;
const advanceUidPast = (id) => {
  if (typeof id === "number" && id >= uidCounter) uidCounter = id + 1;
};

const HEADER_STORAGE_KEY = "quote-app-header";
const GROUPS_STORAGE_KEY = "quote-app-groups";
const EMAIL_STORAGE_KEY = "quote-app-recipient-email";

const initialHeader = {
  company: "",
  productionPlace: "오즈",
  styleNo: "",
  orderQty: "",
};

const loadStoredHeader = () => {
  try {
    const raw = localStorage.getItem(HEADER_STORAGE_KEY);
    if (!raw) return initialHeader;
    const merged = { ...initialHeader, ...JSON.parse(raw) };
    // 생산처 defaults to 오즈 even for saves made before that default existed.
    if (!merged.productionPlace) merged.productionPlace = initialHeader.productionPlace;
    return merged;
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

// Field labels the header block might use, mapped to our current header
// keys. Includes the older 10-field layout (품명/등록번호/성명/COLOR/...)
// so files exported before the header was simplified still import cleanly
// — whatever old fields don't have a home in the current model are just
// dropped, everything else (all item data) is unaffected either way.
const HEADER_LABEL_MAP = {
  업체명: "company",
  상호: "company",
  생산처: "productionPlace",
  스타일넘버: "styleNo",
  "STYLE No": "styleNo",
  발주량: "orderQty",
};

// Normalizes any ExcelJS cell value (plain string/number, rich text with
// mixed fonts, formula result, hyperlink, ...) into plain text.
function cellText(v) {
  if (v == null) return "";
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("").trim();
    if (v.text != null) return String(v.text).trim();
    if (v.result != null) return String(v.result).trim();
    return "";
  }
  return String(v).trim();
}

function findLabelCell(ws, labelText, maxRow = 60, maxCol = 7) {
  for (let r = 1; r <= maxRow; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= maxCol; c++) {
      if (cellText(row.getCell(c).value) === labelText) {
        return { row, col: c };
      }
    }
  }
  return null;
}

async function parseQuoteExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("시트를 찾을 수 없습니다.");

  const header = { ...initialHeader };
  for (let r = 1; r <= 8; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= 7; c++) {
      const mapped = HEADER_LABEL_MAP[cellText(row.getCell(c).value)];
      if (!mapped) continue;
      header[mapped] = cellText(row.getCell(c + 1).value);
    }
  }

  const gubunHit = findLabelCell(ws, "구분");
  if (!gubunHit) throw new Error("견적서 형식을 알아볼 수 없습니다.");
  const itemsStartRow = gubunHit.row.number + 1;

  const productionCostHit = findLabelCell(ws, "생산원가", 200);
  const itemsEndRow = productionCostHit ? productionCostHit.row.number : itemsStartRow;

  let freight = "1000";
  const freightHit = findLabelCell(ws, "운임비", 200);
  if (freightHit) {
    const v = freightHit.row.getCell(7).value;
    if (v != null && v !== "") freight = String(v);
  }

  let marginRate = "15";
  const marginHit = findLabelCell(ws, "업체마진", 200);
  if (marginHit) {
    const v = marginHit.row.getCell(6).value;
    if (typeof v === "number") marginRate = String(Math.round(v * 100));
  }

  // Merged cells report the same value on every row they span, so "is this
  // row's cell non-empty" alone can't tell a new major/sub group apart from
  // a continuation row — only the true top-left (master) cell of a merge
  // marks where a group actually starts.
  const groups = [];
  let currentGroup = null;
  for (let r = itemsStartRow; r < itemsEndRow; r++) {
    const row = ws.getRow(r);
    const majorCell = row.getCell(1);
    const subCell = row.getCell(2);
    const isNewSub = subCell.master === subCell && cellText(subCell.value) !== "";

    if (isNewSub) {
      const isNewMajor = majorCell.master === majorCell && cellText(majorCell.value) !== "";
      currentGroup = {
        id: uid(),
        major: isNewMajor ? cellText(majorCell.value) : groups.length ? groups[groups.length - 1].major : "",
        sub: cellText(subCell.value),
        items: [],
      };
      groups.push(currentGroup);
    }
    if (!currentGroup) continue;

    const name = cellText(row.getCell(3).value);
    const unit = cellText(row.getCell(4).value);
    const price = cellText(row.getCell(5).value);
    const qty = cellText(row.getCell(6).value);
    const hasItem = name !== "" || unit !== "" || price !== "" || qty !== "";
    if (!hasItem) continue;

    const formula = row.getCell(7).formula || "";
    currentGroup.items.push({
      id: uid(),
      name,
      unit,
      price,
      qty,
      amortize: formula.includes("/"),
    });
  }

  if (groups.length === 0) throw new Error("품목 데이터를 찾을 수 없습니다.");

  return { header, groups, freight, marginRate };
}

export default function App({ session }) {
  const [header, setHeader] = useState(loadStoredHeader);
  const [groups, setGroups] = useState(loadStoredGroups);
  const [freight, setFreight] = useState("1000");
  const [marginRate, setMarginRate] = useState("15");
  const [showPreview, setShowPreview] = useState(false);
  const [emailCopied, setEmailCopied] = useState(false);
  const [canShareFile, setCanShareFile] = useState(false);
  const [recipientEmail, setRecipientEmail] = useState(() => {
    try {
      return localStorage.getItem(EMAIL_STORAGE_KEY) || "";
    } catch {
      return "";
    }
  });

  // Cloud save/load (Supabase) — lets the same set of quotes be reached
  // from any device, on top of the per-browser localStorage autosave above.
  const [quoteId, setQuoteId] = useState(null);
  const [quoteName, setQuoteName] = useState("");
  const [showQuoteList, setShowQuoteList] = useState(false);
  const [savedQuotes, setSavedQuotes] = useState([]);
  const [quoteListLoading, setQuoteListLoading] = useState(false);
  const [quoteListError, setQuoteListError] = useState("");
  const [cloudSaving, setCloudSaving] = useState(false);
  const [cloudSaveError, setCloudSaveError] = useState("");
  const [cloudSaved, setCloudSaved] = useState(false);

  useEffect(() => {
    if (!emailCopied) return;
    const t = setTimeout(() => setEmailCopied(false), 4000);
    return () => clearTimeout(t);
  }, [emailCopied]);

  useEffect(() => {
    setCanShareFile(canShareExcel());
  }, []);

  useEffect(() => {
    if (!cloudSaved) return;
    const t = setTimeout(() => setCloudSaved(false), 4000);
    return () => clearTimeout(t);
  }, [cloudSaved]);

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
    setQuoteId(null);
    setQuoteName("");
  };

  const fileInputRef = useRef(null);

  const handleExcelFileSelected = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // allow re-selecting the same file next time
    if (!file) return;

    if (
      !window.confirm(
        "엑셀 파일을 불러오면 현재 화면의 내용을 덮어씁니다. 계속할까요?"
      )
    ) {
      return;
    }

    try {
      const parsed = await parseQuoteExcelFile(file);
      setHeader(parsed.header);
      setGroups(parsed.groups);
      setFreight(parsed.freight);
      setMarginRate(parsed.marginRate);
      setQuoteId(null);
      setQuoteName("");
    } catch (err) {
      window.alert("엑셀 파일을 불러오지 못했습니다: " + err.message);
    }
  };

  const fetchQuoteList = async () => {
    setQuoteListLoading(true);
    setQuoteListError("");
    const { data, error } = await supabase
      .from("quotes")
      .select("id, name, updated_at")
      .order("updated_at", { ascending: false });
    if (error) {
      setQuoteListError("목록을 불러오지 못했습니다: " + error.message);
    } else {
      setSavedQuotes(data || []);
    }
    setQuoteListLoading(false);
  };

  const openQuoteList = () => {
    setShowQuoteList(true);
    fetchQuoteList();
  };

  const saveQuoteToCloud = async () => {
    const name = quoteName.trim() || header.company.trim() || "이름 없는 견적서";
    setCloudSaving(true);
    setCloudSaveError("");

    const payload = {
      name,
      header,
      groups,
      freight,
      margin_rate: marginRate,
    };

    const query = quoteId
      ? supabase.from("quotes").update(payload).eq("id", quoteId).select().single()
      : supabase.from("quotes").insert(payload).select().single();

    const { data, error } = await query;
    if (error) {
      setCloudSaveError("저장하지 못했습니다: " + error.message);
    } else {
      setQuoteId(data.id);
      setQuoteName(data.name);
      setCloudSaved(true);
    }
    setCloudSaving(false);
  };

  const loadQuoteFromCloud = async (id) => {
    setQuoteListError("");
    const { data, error } = await supabase.from("quotes").select("*").eq("id", id).single();
    if (error) {
      setQuoteListError("불러오지 못했습니다: " + error.message);
      return;
    }
    setHeader({ ...initialHeader, ...data.header });
    setGroups(data.groups);
    (data.groups || []).forEach((g) => {
      advanceUidPast(g.id);
      (g.items || []).forEach((it) => advanceUidPast(it.id));
    });
    setFreight(data.freight || "1000");
    setMarginRate(data.margin_rate || "15");
    setQuoteId(data.id);
    setQuoteName(data.name);
    setShowQuoteList(false);
  };

  const deleteQuoteFromCloud = async (id) => {
    const { error } = await supabase.from("quotes").delete().eq("id", id);
    if (error) {
      setQuoteListError("삭제하지 못했습니다: " + error.message);
      return;
    }
    setSavedQuotes((qs) => qs.filter((q) => q.id !== id));
    if (id === quoteId) {
      setQuoteId(null);
      setQuoteName("");
    }
  };

  const groupSubtotal = (g) =>
    g.items.reduce((sum, it) => sum + itemAmount(it, header.orderQty), 0);

  const productionCost = groups.reduce((sum, g) => sum + groupSubtotal(g), 0);
  const freightNum = num(freight);
  const marginAmount = (productionCost + freightNum) * (num(marginRate) / 100);
  const supplyAmount = productionCost + freightNum + marginAmount;

  const buildQuoteWorkbookBuffer = async () => {
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
    ws.mergeCells("B3:D3");
    ws.mergeCells("F3:G3");
    label("A3", "업체명");
    input("B3", header.company);
    label("E3", "생산처");
    input("F3", header.productionPlace);

    ws.mergeCells("B4:D4");
    ws.mergeCells("F4:G4");
    label("A4", "스타일넘버");
    input("B4", header.styleNo);
    label("E4", "발주량");
    input("F4", num(header.orderQty), { numeric: true, numFmt: "#,##0" });

    ws.getRow(3).height = 20;
    ws.getRow(4).height = 20;

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
        const formula = it.amortize ? `E${r}*F${r}/$F$4` : `E${r}*F${r}`;
        setCell(
          `G${r}`,
          { formula, result: Math.round(itemAmount(it, header.orderQty)) },
          { align: "right", numFmt: "#,##0" }
        );
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
    setCell(
      `G${pcRow}`,
      { formula: `SUM(G9:G${lastItemRow})`, result: Math.round(productionCost) },
      { bold: true, fill: LGRAY, align: "right", numFmt: "#,##0" }
    );
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
    setCell(
      `G${mgRow}`,
      { formula: `(G${pcRow}+G${frRow})*F${mgRow}`, result: Math.round(marginAmount) },
      { bold: true, fill: LGRAY, align: "right", numFmt: "#,##0" }
    );
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
      { formula: `G${pcRow}+G${frRow}+G${mgRow}`, result: Math.round(supplyAmount) },
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
    const filename = header.company ? `견적서_${header.company}.xlsx` : "견적서.xlsx";
    return { buffer, filename };
  };

  const exportToExcel = async () => {
    const { buffer, filename } = await buildQuoteWorkbookBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  // Web Share API with a file: on mobile this opens the native share sheet,
  // where Gmail/네이버메일/카카오톡 etc. show up as targets and receive the
  // file already attached — the only way to skip the "download, then attach
  // by hand" step, since mailto: can never carry a file (no browser allows
  // a web page to inject an attachment into an email link).
  const canShareExcel = () => {
    if (!navigator.canShare) return false;
    try {
      const probe = new File([""], "probe.xlsx", {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      return navigator.canShare({ files: [probe] });
    } catch {
      return false;
    }
  };

  const shareQuoteExcel = async () => {
    try {
      const { buffer, filename } = await buildQuoteWorkbookBuffer();
      const file = new File([buffer], filename, {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const { subject, body } = buildEmailContent();
      await navigator.share({ files: [file], title: subject, text: body });
    } catch (err) {
      if (err?.name !== "AbortError") {
        window.alert("공유하지 못했습니다: " + err.message);
      }
    }
  };

  const buildEmailContent = () => {
    const subject = `견적서${header.company ? ` - ${header.company}` : ""}`;
    const body = [
      header.company ? `${header.company} 담당자님께,` : "안녕하세요,",
      "",
      "견적서를 보내드립니다. 첨부된 엑셀 파일을 확인해 주세요.",
      "",
      `- 업체명: ${header.company || "-"}`,
      `- 생산처: ${header.productionPlace || "-"}`,
      `- 스타일넘버: ${header.styleNo || "-"}`,
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
          <div className="flex items-start justify-between gap-3">
            <h1 className="font-serif text-4xl font-bold tracking-wide text-stone-900">
              견 적 서
            </h1>
            <div className="flex flex-col items-end gap-0.5 print:hidden shrink-0 pt-2">
              {session?.user?.email && (
                <span className="text-xs text-stone-400 truncate max-w-[160px]">{session.user.email}</span>
              )}
              <button
                onClick={() => supabase.auth.signOut()}
                className="flex items-center gap-1 text-sm text-stone-500 hover:text-stone-800"
              >
                <LogOut size={14} /> 로그아웃
              </button>
            </div>
          </div>
          <p className="text-sm text-stone-500 mt-1 print:hidden">
            노란 칸에 숫자를 입력하면 금액이 자동으로 계산됩니다.
          </p>
        </div>

        {/* Cloud save/load: keeps a list of quotes in Supabase, reachable from any device */}
        <div className="bg-white border border-stone-300 rounded-lg mb-6 overflow-hidden print:hidden">
          <InfoField
            label="견적서 이름"
            value={quoteName}
            onChange={(e) => setQuoteName(e.target.value)}
            placeholder="예: 모조 앵글부츠 견적"
          />
          <div className="flex flex-col sm:flex-row gap-2 p-3 border-t border-stone-200">
            <button
              onClick={saveQuoteToCloud}
              disabled={cloudSaving}
              className="flex-1 flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-2.5 rounded-lg border border-stone-300 bg-white hover:bg-stone-50 disabled:opacity-50"
            >
              <Cloud size={18} /> {cloudSaving ? "저장 중..." : quoteId ? "저장(업데이트)" : "저장"}
            </button>
            <button
              onClick={openQuoteList}
              className="flex-1 flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-2.5 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
            >
              <FolderOpen size={18} /> 목록 불러오기
            </button>
            {quoteId && (
              <button
                onClick={resetAll}
                className="flex-1 flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-2.5 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
              >
                <FilePlus2 size={18} /> 새 견적서
              </button>
            )}
          </div>
          {cloudSaveError && (
            <p className="text-sm text-red-600 px-3 pb-3">{cloudSaveError}</p>
          )}
          {cloudSaved && (
            <p className="text-sm text-green-700 bg-green-50 px-3 py-2 mx-3 mb-3 rounded-md border border-green-200">
              저장했습니다. 다른 기기에서도 "목록 불러오기"로 열 수 있어요.
            </p>
          )}
        </div>

        {/* Header info: one field per line so every label/input lines up in a column */}
        <div className="bg-white border border-stone-300 rounded-lg mb-6 overflow-hidden divide-y divide-stone-200 print:border-stone-800">
          <InfoField label="업체명" value={header.company} onChange={setHeaderField("company")} />
          <InfoField label="생산처" value={header.productionPlace} onChange={setHeaderField("productionPlace")} />
          <InfoField label="스타일넘버" value={header.styleNo} onChange={setHeaderField("styleNo")} />
          <InfoField label="발주량" value={header.orderQty} onChange={setHeaderField("orderQty")} accent />
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
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xlsm"
            onChange={handleExcelFileSelected}
            className="hidden"
          />
          <button
            onClick={resetAll}
            className="flex items-center justify-center gap-2 text-base text-stone-600 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <RotateCcw size={18} /> 초기값으로
          </button>
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <Upload size={18} /> 엑셀 불러오기
          </button>
          <button
            onClick={() => setShowPreview(true)}
            className="flex items-center justify-center gap-2 text-base text-stone-700 hover:text-stone-900 px-4 py-3 rounded-lg border border-stone-300 bg-white hover:bg-stone-50"
          >
            <Eye size={18} /> 엑셀 미리보기
          </button>
          {canShareFile && (
            <button
              onClick={shareQuoteExcel}
              className="flex items-center justify-center gap-2 text-base font-medium text-white px-4 py-3 rounded-lg bg-blue-600 hover:bg-blue-700"
            >
              <Share2 size={18} /> 파일 첨부해서 공유
            </button>
          )}
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

      {showQuoteList && (
        <QuoteListModal
          quotes={savedQuotes}
          loading={quoteListLoading}
          error={quoteListError}
          currentQuoteId={quoteId}
          onClose={() => setShowQuoteList(false)}
          onLoad={loadQuoteFromCloud}
          onDelete={deleteQuoteFromCloud}
          onRefresh={fetchQuoteList}
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

function QuoteListModal({ quotes, loading, error, currentQuoteId, onClose, onLoad, onDelete, onRefresh }) {
  const formatDate = (iso) => {
    try {
      return new Date(iso).toLocaleString("ko-KR", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      });
    } catch {
      return iso;
    }
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center p-2 sm:p-6 z-50" onClick={onClose}>
      <div
        className="bg-white rounded-lg w-full max-w-lg max-h-full flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3 border-b border-stone-200 shrink-0">
          <h2 className="text-lg font-semibold text-stone-900">저장된 견적서 목록</h2>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-700 p-1" aria-label="닫기">
            <X size={22} />
          </button>
        </div>

        <div className="overflow-auto p-3 sm:p-4">
          {loading && <p className="text-center text-stone-400 py-8">불러오는 중...</p>}
          {!loading && error && (
            <div className="text-center py-8">
              <p className="text-red-600 text-sm mb-3">{error}</p>
              <button
                onClick={onRefresh}
                className="px-4 py-2 rounded-lg border border-stone-300 text-stone-600 hover:bg-stone-50"
              >
                다시 시도
              </button>
            </div>
          )}
          {!loading && !error && quotes.length === 0 && (
            <p className="text-center text-stone-400 py-8">저장된 견적서가 없습니다.</p>
          )}
          {!loading && !error && quotes.length > 0 && (
            <ul className="space-y-2">
              {quotes.map((q) => (
                <li
                  key={q.id}
                  className={
                    "flex items-center justify-between gap-2 border rounded-lg px-3 py-2.5 " +
                    (q.id === currentQuoteId
                      ? "border-amber-400 bg-amber-50"
                      : "border-stone-200 bg-white")
                  }
                >
                  <button
                    onClick={() => onLoad(q.id)}
                    className="flex-1 min-w-0 text-left"
                  >
                    <div className="font-medium text-stone-900 truncate">{q.name}</div>
                    <div className="text-xs text-stone-400">{formatDate(q.updated_at)}</div>
                  </button>
                  <button
                    onClick={() => {
                      if (window.confirm(`"${q.name}"을(를) 삭제할까요? 되돌릴 수 없습니다.`)) {
                        onDelete(q.id);
                      }
                    }}
                    className="text-stone-400 hover:text-red-500 p-2 shrink-0"
                    aria-label="삭제"
                  >
                    <Trash2 size={18} />
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        <div className="flex justify-end px-4 py-3 border-t border-stone-200 shrink-0">
          <button
            onClick={onClose}
            className="px-4 py-2.5 rounded-lg border border-stone-300 text-stone-600 hover:bg-stone-50"
          >
            닫기
          </button>
        </div>
      </div>
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
                <td className={LABEL}>업체명</td>
                <td colSpan={2} className={VALUE}>{header.company}</td>
                <td className={LABEL}>생산처</td>
                <td colSpan={3} className={VALUE}>{header.productionPlace}</td>
              </tr>
              <tr>
                <td className={LABEL}>스타일넘버</td>
                <td colSpan={2} className={VALUE}>{header.styleNo}</td>
                <td className={LABEL}>발주량</td>
                <td colSpan={3} className={VALUE + " text-right"}>{header.orderQty}</td>
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
