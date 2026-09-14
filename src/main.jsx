import React from "react";
import ReactDOM from "react-dom/client";
import { HashRouter, Routes, Route } from "react-router-dom";
import App from "./App.jsx";
import AdminPage from "./AdminPage.jsx";
import AuthGate from "./AuthGate.jsx";
import "./index.css";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <HashRouter>
      <AuthGate>
        {(session) => (
          <Routes>
            <Route path="/" element={<App session={session} />} />
            <Route path="/admin" element={<AdminPage session={session} />} />
          </Routes>
        )}
      </AuthGate>
    </HashRouter>
  </React.StrictMode>
);
