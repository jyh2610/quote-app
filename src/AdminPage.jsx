import { useEffect, useMemo, useState } from "react";
import { Link } from "react-router-dom";
import { ArrowLeft, LogOut, Plus, Trash2, Copy, RefreshCw } from "lucide-react";
import { supabase } from "./supabaseClient";
import { won, computeQuoteTotals } from "./quoteMath";

export default function AdminPage({ session }) {
  const [tab, setTab] = useState("contacts");

  return (
    <div className="min-h-screen bg-stone-100 text-stone-800 text-base pb-16">
      <div className="max-w-3xl mx-auto px-4 py-8">
        <div className="border-b-4 border-stone-800 pb-3 mb-6">
          <div className="flex items-start justify-between gap-3">
            <div>
              <h1 className="font-serif text-3xl font-bold tracking-wide text-stone-900">관 리</h1>
              <Link
                to="/"
                className="inline-flex items-center gap-1 text-sm text-stone-500 hover:text-stone-800 mt-2"
              >
                <ArrowLeft size={14} /> 견적서로 돌아가기
              </Link>
            </div>
            <div className="flex flex-col items-end gap-0.5 shrink-0 pt-2">
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
        </div>

        <div className="flex gap-2 mb-6">
          <TabButton active={tab === "contacts"} onClick={() => setTab("contacts")}>
            주소록
          </TabButton>
          <TabButton active={tab === "stats"} onClick={() => setTab("stats")}>
            통계
          </TabButton>
        </div>

        {tab === "contacts" ? <ContactsTab /> : <StatsTab />}
      </div>
    </div>
  );
}

function TabButton({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      className={
        "flex-1 text-base font-medium px-4 py-3 rounded-lg border " +
        (active
          ? "bg-stone-800 text-white border-stone-800"
          : "bg-white text-stone-600 border-stone-300 hover:bg-stone-50")
      }
    >
      {children}
    </button>
  );
}

function ContactsTab() {
  const [contacts, setContacts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [label, setLabel] = useState("");
  const [email, setEmail] = useState("");
  const [saving, setSaving] = useState(false);
  const [copiedId, setCopiedId] = useState(null);

  const load = async () => {
    setLoading(true);
    setError("");
    const { data, error } = await supabase
      .from("contacts")
      .select("*")
      .order("created_at", { ascending: false });
    if (error) {
      setError("주소록을 불러오지 못했습니다: " + error.message);
    } else {
      setContacts(data || []);
    }
    setLoading(false);
  };

  useEffect(() => {
    load();
  }, []);

  const addContact = async (e) => {
    e.preventDefault();
    const trimmedEmail = email.trim();
    if (!trimmedEmail) return;
    setSaving(true);
    setError("");
    const { error } = await supabase
      .from("contacts")
      .insert({ label: label.trim(), email: trimmedEmail });
    if (error) {
      setError("추가하지 못했습니다: " + error.message);
    } else {
      setLabel("");
      setEmail("");
      await load();
    }
    setSaving(false);
  };

  const removeContact = async (id) => {
    if (!window.confirm("이 주소를 삭제할까요?")) return;
    const { error } = await supabase.from("contacts").delete().eq("id", id);
    if (error) {
      setError("삭제하지 못했습니다: " + error.message);
      return;
    }
    setContacts((cs) => cs.filter((c) => c.id !== id));
  };

  const copyEmail = async (c) => {
    try {
      await navigator.clipboard.writeText(c.email);
      setCopiedId(c.id);
      setTimeout(() => setCopiedId((id) => (id === c.id ? null : id)), 1500);
    } catch {
      // Clipboard API can fail silently (no focus, no permission) — not
      // worth surfacing an error for a convenience action.
    }
  };

  return (
    <div>
      <form
        onSubmit={addContact}
        className="bg-white border border-stone-300 rounded-lg p-4 mb-6 flex flex-col sm:flex-row gap-3"
      >
        <input
          value={label}
          onChange={(e) => setLabel(e.target.value)}
          placeholder="이름/메모 (선택)"
          className="flex-1 border border-stone-300 rounded-md px-3 py-3 text-base outline-none focus:border-stone-500"
        />
        <input
          type="email"
          required
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder="이메일 주소 (필수)"
          className="flex-[2] border border-stone-300 rounded-md px-3 py-3 text-base outline-none focus:border-stone-500"
        />
        <button
          type="submit"
          disabled={saving}
          className="flex items-center justify-center gap-2 text-base font-medium bg-stone-800 text-white px-5 py-3 rounded-lg hover:bg-stone-700 disabled:opacity-50"
        >
          <Plus size={18} /> 추가
        </button>
      </form>

      {error && (
        <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-md px-3 py-2 mb-4">
          {error}
        </p>
      )}

      <div className="bg-white border border-stone-300 rounded-lg overflow-hidden">
        {loading ? (
          <p className="text-center text-stone-400 py-10">불러오는 중...</p>
        ) : contacts.length === 0 ? (
          <p className="text-center text-stone-400 py-10">저장된 이메일 주소가 없습니다.</p>
        ) : (
          <ul className="divide-y divide-stone-200">
            {contacts.map((c) => (
              <li key={c.id} className="flex items-center justify-between gap-3 px-4 py-3">
                <div className="min-w-0">
                  {c.label && (
                    <p className="text-sm text-stone-500 truncate">{c.label}</p>
                  )}
                  <p className="text-base text-stone-900 truncate">{c.email}</p>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button
                    onClick={() => copyEmail(c)}
                    className="flex items-center gap-1 text-sm text-stone-500 hover:text-stone-800 px-2 py-2"
                  >
                    <Copy size={16} /> {copiedId === c.id ? "복사됨" : "복사"}
                  </button>
                  <button
                    onClick={() => removeContact(c.id)}
                    className="text-stone-400 hover:text-red-600 p-2"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

function StatsTab() {
  const [quotes, setQuotes] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = async () => {
    setLoading(true);
    setError("");
    const { data, error } = await supabase
      .from("quotes")
      .select("id, name, header, groups, freight, margin_rate, created_at")
      .order("created_at", { ascending: false });
    if (error) {
      setError("통계를 불러오지 못했습니다: " + error.message);
    } else {
      setQuotes(data || []);
    }
    setLoading(false);
  };

  useEffect(() => {
    load();
  }, []);

  const stats = useMemo(() => {
    const now = new Date();
    const thisMonthKey = `${now.getFullYear()}-${now.getMonth()}`;

    let totalSupply = 0;
    let monthCount = 0;
    let monthSupply = 0;
    const byPlace = new Map();
    const rows = [];

    for (const q of quotes) {
      const { supplyAmount } = computeQuoteTotals(
        q.header || {},
        q.groups || [],
        q.freight,
        q.margin_rate
      );
      totalSupply += supplyAmount;

      const created = q.created_at ? new Date(q.created_at) : null;
      const key = created ? `${created.getFullYear()}-${created.getMonth()}` : null;
      if (key === thisMonthKey) {
        monthCount += 1;
        monthSupply += supplyAmount;
      }

      const place = (q.header?.productionPlace || "").trim() || "미지정";
      byPlace.set(place, (byPlace.get(place) || 0) + 1);

      rows.push({
        id: q.id,
        name: q.name,
        company: q.header?.company || "",
        place,
        created,
        supplyAmount,
      });
    }

    const byPlaceList = [...byPlace.entries()].sort((a, b) => b[1] - a[1]);

    return {
      total: quotes.length,
      totalSupply,
      monthCount,
      monthSupply,
      byPlaceList,
      rows: rows.slice(0, 10),
    };
  }, [quotes]);

  if (loading) {
    return <p className="text-center text-stone-400 py-10">불러오는 중...</p>;
  }

  if (error) {
    return (
      <p className="text-sm text-red-600 bg-red-50 border border-red-200 rounded-md px-3 py-2">
        {error}
      </p>
    );
  }

  return (
    <div>
      <div className="flex justify-end mb-3">
        <button
          onClick={load}
          className="flex items-center gap-1 text-sm text-stone-500 hover:text-stone-800"
        >
          <RefreshCw size={14} /> 새로고침
        </button>
      </div>

      <div className="grid grid-cols-2 gap-3 mb-6">
        <StatCard label="전체 견적서" value={`${stats.total}건`} />
        <StatCard label="전체 공급가액 합계" value={`${won(stats.totalSupply)}원`} />
        <StatCard label="이번 달 견적서" value={`${stats.monthCount}건`} />
        <StatCard label="이번 달 공급가액 합계" value={`${won(stats.monthSupply)}원`} />
      </div>

      <div className="bg-white border border-stone-300 rounded-lg p-4 mb-6">
        <h2 className="text-base font-semibold text-stone-800 mb-3">생산처별 견적서 수</h2>
        {stats.byPlaceList.length === 0 ? (
          <p className="text-sm text-stone-400">데이터가 없습니다.</p>
        ) : (
          <ul className="space-y-2">
            {stats.byPlaceList.map(([place, count]) => (
              <li key={place} className="flex items-center justify-between text-sm">
                <span className="text-stone-700">{place}</span>
                <span className="text-stone-500">{count}건</span>
              </li>
            ))}
          </ul>
        )}
      </div>

      <div className="bg-white border border-stone-300 rounded-lg overflow-hidden">
        <h2 className="text-base font-semibold text-stone-800 px-4 pt-4">최근 견적서</h2>
        {stats.rows.length === 0 ? (
          <p className="text-sm text-stone-400 px-4 py-6">저장된 견적서가 없습니다.</p>
        ) : (
          <ul className="divide-y divide-stone-200 mt-2">
            {stats.rows.map((r) => (
              <li key={r.id} className="flex items-center justify-between gap-3 px-4 py-3">
                <div className="min-w-0">
                  <p className="text-base text-stone-900 truncate">{r.name}</p>
                  <p className="text-sm text-stone-500 truncate">
                    {r.company} · {r.place}
                    {r.created ? ` · ${r.created.toLocaleDateString("ko-KR")}` : ""}
                  </p>
                </div>
                <span className="text-base tabular-nums text-stone-800 shrink-0">
                  {won(r.supplyAmount)}원
                </span>
              </li>
            ))}
          </ul>
        )}
      </div>
    </div>
  );
}

function StatCard({ label, value }) {
  return (
    <div className="bg-white border border-stone-300 rounded-lg p-4">
      <p className="text-sm text-stone-500 mb-1">{label}</p>
      <p className="text-xl font-semibold tabular-nums text-stone-900">{value}</p>
    </div>
  );
}
