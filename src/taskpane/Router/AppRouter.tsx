import React, { useState } from "react";
import { Routes, Route, Navigate } from "react-router-dom";
import Home from "../components/pages/Home";
import Register from "../components/pages/Register";
import ReportError from "../components/pages/ReportError";

function AppRoutes() {
  const [token, setToken] = useState(() => localStorage.getItem("token"));

  return (
    <Routes>
      <Route
        path="/register"
        element={<Register onConnected={() => setToken(localStorage.getItem("token"))} />}
      />
      <Route
        path="/home"
        element={token ? <Home /> : <Navigate to="/register" replace />}
      />
      <Route
        path="*"
        element={<Navigate to={token ? "/home" : "/register"} replace />}
      />
      <Route path="/report" element={<ReportError />} />
    </Routes>
  );
} 

export default AppRoutes;