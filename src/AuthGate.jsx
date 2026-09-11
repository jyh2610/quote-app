import { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";

// Gates the app behind Supabase Auth. There's no sign-up form on purpose —
// accounts are created directly in the Supabase dashboard
// (Authentication -> Users -> Add user), not through this page.
export default function AuthGate({ children }) {
  const [session, setSession] = useState(undefined); // undefined = still checking, null = signed out

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-stone-100 text-stone-500 text-base">
        불러오는 중...
      </div>
    );
  }

  if (!session) {
    return <LoginScreen />;
  }

  return children(session);
}

function LoginScreen() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) {
      setError("로그인에 실패했습니다. 이메일과 비밀번호를 확인해 주세요.");
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-stone-100 px-4">
      <form
        onSubmit={handleLogin}
        className="bg-white border border-stone-300 rounded-lg p-6 sm:p-8 w-full max-w-sm"
      >
        <h1 className="font-serif text-3xl font-bold text-stone-900 mb-1">견 적 서</h1>
        <p className="text-sm text-stone-500 mb-6">로그인 후 이용할 수 있습니다.</p>

        <label className="block text-sm text-stone-600 mb-1" htmlFor="login-email">
          이메일
        </label>
        <input
          id="login-email"
          type="email"
          required
          autoComplete="username"
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          className="w-full border border-stone-300 rounded-md px-3 py-3 mb-4 text-base outline-none focus:border-stone-500"
        />

        <label className="block text-sm text-stone-600 mb-1" htmlFor="login-password">
          비밀번호
        </label>
        <input
          id="login-password"
          type="password"
          required
          autoComplete="current-password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          className="w-full border border-stone-300 rounded-md px-3 py-3 mb-4 text-base outline-none focus:border-stone-500"
        />

        {error && <p className="text-sm text-red-600 mb-4">{error}</p>}

        <button
          type="submit"
          disabled={loading}
          className="w-full bg-stone-800 text-white rounded-lg py-3 text-base font-medium hover:bg-stone-700 disabled:opacity-50"
        >
          {loading ? "로그인 중..." : "로그인"}
        </button>
      </form>
    </div>
  );
}
