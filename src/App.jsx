import { useEffect, useState } from "react";
import { Plus, Trash2, FileSpreadsheet, RotateCcw } from "lucide-react";
import * as XLSX from "xlsx";

let uidCounter = 1;
const uid = () => uidCounter++;
const advanceUidPast = (id) => {
  if (typeof id === "number" && id >= uidCounter) uidCounter = id + 1;
};

const HEADER_STORAGE_KEY = "quote-app-header";
const GROUPS_STORAGE_KEY = "quote-app-groups";

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

export default function App() {
  const [header, setHeader] = useState(loadStoredHeader);
  const [groups, setGroups] = useState(loadStoredGroups);
  const [freight, setFreight] = useState("1000");
  const [marginRate, setMarginRate] = useState("15");

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
          : { ...g, items: [...g.items, seedItem("", "족", "", "1")] }
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

  const exportToExcel = () => {
    const rows = [
      ["견 적 서"],
      [],
      ["품명", header.productName, "", "등록번호", header.regNo],
      ["STYLE No", header.styleNo, "", "상호", header.company],
      ["발주량", header.orderQty, "", "성명", header.ceoName],
      ["COLOR", header.color, "", "업태", header.bizType],
      ["종목", header.category, "", "사업장주소", header.address],
      [],
    ];

    groups.forEach((g) => {
      rows.push([`${g.major} / ${g.sub}`]);
      rows.push(["품목", "소재", "단가", "소요량", "금액"]);
      g.items.forEach((it) => {
        rows.push([
          it.name,
          it.unit,
          num(it.price),
          num(it.qty),
          Math.round(itemAmount(it, header.orderQty)),
        ]);
      });
      rows.push(["", "", "", "소계", Math.round(groupSubtotal(g))]);
      rows.push([]);
    });

    rows.push(["생산원가", "", "", "", Math.round(productionCost)]);
    rows.push(["운임비", "", "", "", freightNum]);
    rows.push([`업체마진 (${marginRate}%)`, "", "", "", Math.round(marginAmount)]);
    rows.push(["공급가액", "", "", "", Math.round(supplyAmount)]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "견적서");

    const filename = header.productName
      ? `견적서_${header.productName}.xlsx`
      : "견적서.xlsx";
    XLSX.writeFile(wb, filename);
  };

  return (
    <div className="min-h-screen bg-stone-100 text-stone-800 pb-24 print:bg-white print:pb-0">
      <div className="max-w-3xl mx-auto px-4 py-8 print:p-0 print:max-w-none">
        {/* Title */}
        <div className="flex items-baseline justify-between border-b-4 border-stone-800 pb-3 mb-6">
          <h1 className="font-serif text-3xl tracking-wide text-stone-900">
            견 적 서
          </h1>
          <span className="text-xs text-stone-400 print:hidden">
            숫자를 바꾸면 자동 계산됩니다
          </span>
        </div>

        {/* Header info */}
        <div className="grid grid-cols-2 gap-x-6 gap-y-3 bg-white border border-stone-300 rounded-lg p-4 mb-6 print:border-stone-800">
          <Field label="품명" value={header.productName} onChange={setHeaderField("productName")} />
          <Field label="등록번호" value={header.regNo} onChange={setHeaderField("regNo")} />
          <Field label="STYLE No" value={header.styleNo} onChange={setHeaderField("styleNo")} />
          <Field label="상호" value={header.company} onChange={setHeaderField("company")} />
          <Field label="발주량" value={header.orderQty} onChange={setHeaderField("orderQty")} accent />
          <Field label="성명" value={header.ceoName} onChange={setHeaderField("ceoName")} />
          <Field label="COLOR" value={header.color} onChange={setHeaderField("color")} />
          <Field label="업태" value={header.bizType} onChange={setHeaderField("bizType")} />
          <Field label="종목" value={header.category} onChange={setHeaderField("category")} />
          <Field label="사업장주소" value={header.address} onChange={setHeaderField("address")} wide />
        </div>

        {/* Item groups */}
        <div className="space-y-6">
          {groups.map((g) => (
            <div key={g.id} className="bg-white border border-stone-300 rounded-lg overflow-hidden print:border-stone-800">
              <div className="flex items-center justify-between bg-stone-800 text-stone-50 px-4 py-2">
                <div className="font-serif text-sm tracking-wide">
                  {g.major} <span className="text-stone-300 mx-1">/</span> {g.sub}
                </div>
                <div className="text-sm font-medium tabular-nums">
                  {won(groupSubtotal(g))}원
                </div>
              </div>

              {/* column labels, desktop only */}
              <div className="hidden sm:grid grid-cols-12 gap-2 px-4 pt-3 text-xs text-stone-400">
                <div className="col-span-4">품목</div>
                <div className="col-span-2">소재</div>
                <div className="col-span-2 text-right">단가</div>
                <div className="col-span-2 text-right">소요량</div>
                <div className="col-span-2 text-right">금액</div>
              </div>

              <div className="px-4 pb-3">
                {g.items.map((it) => (
                  <div
                    key={it.id}
                    className="grid grid-cols-2 sm:grid-cols-12 gap-2 py-2 border-b border-stone-100 last:border-b-0 items-center"
                  >
                    <div className="col-span-2 sm:col-span-4">
                      <MiniLabel text="품목" />
                      <input
                        className="w-full bg-transparent border-b border-stone-200 focus:border-stone-500 outline-none py-1 text-sm print:border-none"
                        value={it.name}
                        placeholder="품목명"
                        onChange={(e) => updateItem(g.id, it.id, "name", e.target.value)}
                      />
                    </div>
                    <div>
                      <MiniLabel text="소재" />
                      <input
                        className="w-full bg-transparent border-b border-stone-200 focus:border-stone-500 outline-none py-1 text-sm print:border-none"
                        value={it.unit}
                        placeholder="단위"
                        onChange={(e) => updateItem(g.id, it.id, "unit", e.target.value)}
                      />
                    </div>
                    <div>
                      <MiniLabel text="단가" />
                      <input
                        className="w-full bg-amber-50 border-b border-amber-200 focus:border-amber-500 outline-none py-1 text-sm text-right tabular-nums print:bg-transparent print:border-none"
                        inputMode="decimal"
                        value={it.price}
                        onChange={(e) => updateItem(g.id, it.id, "price", e.target.value)}
                      />
                    </div>
                    <div className="flex items-center gap-1">
                      <div className="flex-1">
                        <MiniLabel text="소요량" />
                        <input
                          className="w-full bg-amber-50 border-b border-amber-200 focus:border-amber-500 outline-none py-1 text-sm text-right tabular-nums print:bg-transparent print:border-none"
                          inputMode="decimal"
                          value={it.qty}
                          onChange={(e) => updateItem(g.id, it.id, "qty", e.target.value)}
                        />
                      </div>
                    </div>
                    <div className="col-span-2 sm:col-span-2 flex items-center justify-between sm:justify-end gap-2">
                      <span className="text-sm font-medium tabular-nums">
                        {won(itemAmount(it, header.orderQty))}
                      </span>
                      <button
                        onClick={() => removeItem(g.id, it.id)}
                        className="text-stone-300 hover:text-red-500 print:hidden"
                        aria-label="행 삭제"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                    <label className="col-span-2 sm:col-span-12 flex items-center gap-1.5 text-xs text-stone-400 print:hidden -mt-1">
                      <input
                        type="checkbox"
                        checked={it.amortize}
                        onChange={(e) => updateItem(g.id, it.id, "amortize", e.target.checked)}
                        className="accent-amber-700"
                      />
                      발주량({header.orderQty || 0})으로 나눠서 계산 (금형비 등 1회성 비용)
                    </label>
                  </div>
                ))}
              </div>

              <button
                onClick={() => addItem(g.id)}
                className="w-full flex items-center justify-center gap-1 text-xs text-stone-500 hover:text-stone-800 hover:bg-stone-50 py-2 border-t border-stone-100 print:hidden"
              >
                <Plus size={14} /> {g.sub}에 항목 추가
              </button>
            </div>
          ))}
        </div>

        {/* Totals */}
        <div className="mt-6 bg-white border border-stone-300 rounded-lg p-4 print:border-stone-800">
          <TotalRow label="생산원가" value={won(productionCost)} />
          <div className="flex items-center justify-between py-2 border-t border-stone-100">
            <span className="text-sm text-stone-600">운임비</span>
            <input
              className="w-28 bg-amber-50 border-b border-amber-200 focus:border-amber-500 outline-none py-1 text-sm text-right tabular-nums print:bg-transparent print:border-none"
              inputMode="decimal"
              value={freight}
              onChange={(e) => setFreight(e.target.value)}
            />
          </div>
          <div className="flex items-center justify-between py-2 border-t border-stone-100">
            <span className="text-sm text-stone-600">업체마진</span>
            <div className="flex items-center gap-3">
              <div className="flex items-center gap-1">
                <input
                  className="w-14 bg-amber-50 border-b border-amber-200 focus:border-amber-500 outline-none py-1 text-sm text-right tabular-nums print:bg-transparent print:border-none"
                  inputMode="decimal"
                  value={marginRate}
                  onChange={(e) => setMarginRate(e.target.value)}
                />
                <span className="text-sm text-stone-400">%</span>
              </div>
              <span className="text-sm font-medium tabular-nums w-24 text-right">
                {won(marginAmount)}
              </span>
            </div>
          </div>
          <div className="flex items-center justify-between pt-3 mt-1 border-t-2 border-stone-800">
            <span className="font-serif text-base text-stone-900">공급가액</span>
            <span className="font-serif text-xl font-semibold tabular-nums text-stone-900">
              {won(supplyAmount)}원
            </span>
          </div>
        </div>

        <div className="flex justify-end gap-3 mt-4 print:hidden">
          <button
            onClick={resetAll}
            className="flex items-center gap-1.5 text-sm text-stone-500 hover:text-stone-800 px-3 py-2"
          >
            <RotateCcw size={14} /> 초기값으로
          </button>
          <button
            onClick={exportToExcel}
            className="flex items-center gap-1.5 text-sm bg-stone-800 text-white px-4 py-2 rounded-lg hover:bg-stone-700"
          >
            <FileSpreadsheet size={14} /> 엑셀로 저장
          </button>
        </div>
      </div>

      {/* Sticky mobile summary */}
      <div className="fixed bottom-0 inset-x-0 bg-stone-800 text-stone-50 px-4 py-3 flex items-center justify-between sm:hidden print:hidden">
        <span className="text-xs text-stone-300">공급가액</span>
        <span className="font-serif text-lg font-semibold tabular-nums">
          {won(supplyAmount)}원
        </span>
      </div>
    </div>
  );
}

function Field({ label, value, onChange, wide, accent }) {
  return (
    <div className={wide ? "col-span-2" : ""}>
      <div className="text-xs text-stone-400 mb-0.5">{label}</div>
      <input
        className={
          "w-full outline-none border-b py-1 text-sm " +
          (accent
            ? "bg-amber-50 border-amber-200 focus:border-amber-500"
            : "bg-transparent border-stone-200 focus:border-stone-500") +
          " print:bg-transparent print:border-none"
        }
        value={value}
        onChange={onChange}
      />
    </div>
  );
}

function MiniLabel({ text }) {
  return <div className="sm:hidden text-xs text-stone-400">{text}</div>;
}

function TotalRow({ label, value }) {
  return (
    <div className="flex items-center justify-between py-2">
      <span className="text-sm text-stone-600">{label}</span>
      <span className="text-sm font-medium tabular-nums">{value}</span>
    </div>
  );
}
