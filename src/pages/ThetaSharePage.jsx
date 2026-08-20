import React, { useEffect, useRef, useState } from "react";
import { useNavigate, useParams } from "react-router-dom";
import toast from "react-hot-toast";
import { Loader2, Link2, Download } from "lucide-react";
import * as XLSX from "xlsx";
import SpreadsheetEditor from "../components/SpreadsheetEditor";
import {
  LOGIN_URL,
  thetaFileService,
  thetaFileSharePath,
  thetaFileShareUrl,
} from "../services/api";
import useStore from "../store/useStore";

const gridToXlsxFile = (gridData, fileName = "Theta Sheets.xlsx") => {
  const sheets = gridData?.sheets || [];
  const wb = XLSX.utils.book_new();
  sheets.forEach((s) => {
    const aoa = [s.headers || [], ...(s.rows || [])];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, (s.name || "Sheet1").slice(0, 31));
  });
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return new File([wbout], fileName, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
};

const findHeaderRowIndex = (raw, scanRows = 10) => {
  let bestIdx = 0;
  let bestCount = -1;
  for (let i = 0; i < Math.min(scanRows, raw.length); i++) {
    const count = (raw[i] || []).filter(
      (c) => String(c ?? "").trim() !== "",
    ).length;
    if (count > bestCount) {
      bestCount = count;
      bestIdx = i;
    }
  }
  return bestIdx;
};

const parseSheetWithHeaderDetection = (ws, name) => {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  const headerIdx = findHeaderRowIndex(raw);
  const headers = (raw[headerIdx] || []).map((h) => String(h).trim());
  const rows = raw
    .slice(headerIdx + 1)
    .filter((row) => row.some((c) => String(c ?? "").trim() !== ""));
  return { name, headers, rows };
};

const ACCESS_DENIED = "You need access";

const ThetaSharePage = () => {
  const navigate = useNavigate();
  const { fileId, linkToken, shareToken } = useParams();
  const { isAuthenticated } = useStore();
  const editorRef = useRef(null);

  const [status, setStatus] = useState("loading");
  const [message, setMessage] = useState("");
  const [fileEntry, setFileEntry] = useState(null);
  const [sheets, setSheets] = useState([]);
  const [saving, setSaving] = useState(false);

  const resolvedFileId = fileId || fileEntry?.id;
  const resolvedToken = linkToken || shareToken || fileEntry?.link_token;
  const isAnonymous = !isAuthenticated;
  const needsLoginToSave = Boolean(
    isAnonymous &&
      (fileEntry?.requires_login_to_save ||
        fileEntry?.link_permission === "editor"),
  );
  const accessRole = String(fileEntry?.access_role || "").toLowerCase();
  const canEdit =
    !isAnonymous &&
    (fileEntry?.can_edit === true ||
      accessRole === "owner" ||
      accessRole === "editor" ||
      fileEntry?.is_owner === true);
  const readOnly = !canEdit;
  const roleLabel =
    accessRole === "owner" || fileEntry?.is_owner
      ? "Owner"
      : accessRole === "editor"
        ? "Editor"
        : needsLoginToSave
          ? "View only — sign in to save"
          : "View only";

  const goLogin = () => {
    const next =
      fileId && linkToken
        ? thetaFileSharePath(fileId, linkToken)
        : `${window.location.pathname}${window.location.search}`;
    const login = LOGIN_URL.includes("?")
      ? `${LOGIN_URL}&next=${encodeURIComponent(next)}`
      : `${LOGIN_URL}?next=${encodeURIComponent(next)}`;
    window.location.href = login;
  };

  const copyLink = async () => {
    if (!resolvedFileId || !resolvedToken) {
      toast.error("Could not copy link.");
      return;
    }
    try {
      await navigator.clipboard.writeText(
        thetaFileShareUrl(resolvedFileId, resolvedToken),
      );
      toast.success("Link copied");
    } catch {
      toast.error("Could not copy link.");
    }
  };

  useEffect(() => {
    let cancelled = false;

    const openDenied = (text = ACCESS_DENIED) => {
      setStatus("denied");
      setMessage(text);
      toast.error(text);
    };

    const loadWorkbook = async (entry, token) => {
      const blob = await thetaFileService.downloadShareBlob(entry.id, token);
      const buf = await blob.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      return wb.SheetNames.map((name) =>
        parseSheetWithHeaderDetection(wb.Sheets[name], name),
      );
    };

    (async () => {
      try {
        let entry;
        let token = (linkToken || shareToken || "").trim();
        if (fileId && linkToken) {
          entry = await thetaFileService.resolveShare(fileId, linkToken);
          token = linkToken;
        } else if (shareToken) {
          entry = await thetaFileService.getByLink(shareToken);
          token = shareToken;
          if (entry?.id && entry?.link_token) {
            navigate(thetaFileSharePath(entry.id, entry.link_token), {
              replace: true,
            });
            return;
          }
        } else {
          openDenied("This link is invalid or has expired.");
          return;
        }
        if (cancelled) return;
        const parsed = await loadWorkbook(entry, token);
        if (cancelled) return;
        setFileEntry(entry);
        setSheets(parsed);
        setStatus("ready");
      } catch (err) {
        if (cancelled) return;
        const statusCode = err?.response?.status;
        const data = err?.response?.data || {};
        if (statusCode === 401 && data.requires_login) {
          goLogin();
          return;
        }
        if (statusCode === 404) {
          openDenied("This link is invalid or has expired.");
          return;
        }
        if (statusCode === 403 || statusCode === 401) {
          openDenied(ACCESS_DENIED);
          return;
        }
        openDenied(data.error || "This link is invalid or has expired.");
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [fileId, linkToken, shareToken, navigate]);

  const handleDownload = async () => {
    if (!resolvedFileId || !resolvedToken) return;
    try {
      const blob = await thetaFileService.downloadShareBlob(
        resolvedFileId,
        resolvedToken,
      );
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = fileEntry?.filename || "Theta Sheets.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      toast.error(err?.response?.data?.error || "Could not download this file.");
    }
  };

  const handleSave = async () => {
    if (!resolvedFileId || !resolvedToken) return;
    if (isAnonymous || needsLoginToSave) {
      goLogin();
      return;
    }
    if (readOnly) return;
    const liveGrid = editorRef.current?.getGrid();
    const liveSheets = liveGrid?.sheets?.length ? liveGrid.sheets : sheets;
    setSaving(true);
    try {
      const file = gridToXlsxFile(
        { name: fileEntry?.filename || "Theta Sheets", sheets: liveSheets },
        fileEntry?.filename || "Theta Sheets.xlsx",
      );
      await thetaFileService.replaceShare(resolvedFileId, resolvedToken, file);
      setSheets(liveSheets);
      toast.success("Saved.");
    } catch (err) {
      const statusCode = err?.response?.status;
      const data = err?.response?.data || {};
      if (statusCode === 401 && data.requires_login) {
        goLogin();
        return;
      }
      toast.error(data.error || "Could not save this file.");
    } finally {
      setSaving(false);
    }
  };

  if (status === "loading") {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
          fontFamily: "sans-serif",
          color: "#64748b",
        }}
      >
        <div style={{ textAlign: "center" }}>
          <Loader2
            size={22}
            style={{
              marginBottom: 8,
              animation: "spin 1s linear infinite",
            }}
          />
          <div>Opening shared file…</div>
          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
        </div>
      </div>
    );
  }

  if (status === "denied") {
    return (
      <div
        style={{
          minHeight: "100vh",
          display: "grid",
          placeItems: "center",
          fontFamily: "sans-serif",
          padding: 24,
        }}
      >
        <div style={{ maxWidth: 420, textAlign: "center" }}>
          <h1 style={{ fontSize: 20, margin: "0 0 8px", color: "#0f172a" }}>
            {message || ACCESS_DENIED}
          </h1>
          <p style={{ color: "#64748b", margin: "0 0 20px" }}>
            {String(message || "").toLowerCase().includes("invalid")
              ? "Check that you copied the full link."
              : "This file is not available with your current access."}
          </p>
          {!isAuthenticated &&
            !String(message || "").toLowerCase().includes("invalid") && (
            <button
              type="button"
              onClick={() => {
                const next = `${window.location.pathname}${window.location.search}`;
                window.location.href = `${LOGIN_URL}?next=${encodeURIComponent(next)}`;
              }}
              style={{
                padding: "9px 16px",
                border: "none",
                borderRadius: 8,
                background: "#059669",
                color: "#fff",
                fontWeight: 600,
                cursor: "pointer",
              }}
            >
              Sign in
            </button>
          )}
        </div>
      </div>
    );
  }

  return (
    <div
      style={{
        height: "100vh",
        display: "flex",
        flexDirection: "column",
        background: "#f8fafc",
        fontFamily: "sans-serif",
      }}
    >
      <header
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
          padding: "12px 20px",
          background: "#fff",
          borderBottom: "1px solid #e2e8f0",
        }}
      >
        <div style={{ minWidth: 0 }}>
          <div
            style={{
              fontSize: 15,
              fontWeight: 700,
              color: "#0f172a",
              overflow: "hidden",
              textOverflow: "ellipsis",
              whiteSpace: "nowrap",
            }}
          >
            {fileEntry?.filename || "Shared file"}
          </div>
          <div style={{ fontSize: 12, color: "#64748b", marginTop: 2 }}>
            {roleLabel}
          </div>
        </div>
        <button
          type="button"
          onClick={copyLink}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            padding: "8px 12px",
            border: "1px solid #cbd5e1",
            borderRadius: 7,
            background: "#fff",
            color: "#334155",
            fontWeight: 600,
            cursor: "pointer",
            flexShrink: 0,
          }}
        >
          <Link2 size={14} /> Copy link
        </button>
      </header>
      <div style={{ flex: 1, minHeight: 0, padding: 16 }}>
        <div
          style={{
            height: "100%",
            background: "#fff",
            border: "1px solid #e2e8f0",
            borderRadius: 8,
            overflow: "hidden",
          }}
        >
          <SpreadsheetEditor
            ref={editorRef}
            key={fileEntry?.id || "share"}
            initialData={{
              name: fileEntry?.filename || "Theta Sheets",
              sheets,
            }}
            hideToolbar
            readOnly={readOnly}
            onCopyLink={copyLink}
            height="100%"
          />
        </div>
      </div>
      <footer
        style={{
          display: "flex",
          justifyContent: "flex-end",
          alignItems: "center",
          gap: 10,
          padding: "12px 20px",
          background: "#fff",
          borderTop: "1px solid #e2e8f0",
        }}
      >
        <button
          type="button"
          onClick={handleDownload}
          style={{
            display: "inline-flex",
            alignItems: "center",
            gap: 6,
            padding: "9px 14px",
            border: "1px solid #cbd5e1",
            borderRadius: 8,
            background: "#fff",
            color: "#334155",
            fontWeight: 600,
            cursor: "pointer",
          }}
        >
          <Download size={14} /> Download
        </button>
        {needsLoginToSave ? (
          <button
            type="button"
            onClick={goLogin}
            style={{
              padding: "9px 16px",
              border: "none",
              borderRadius: 8,
              background: "#059669",
              color: "#fff",
              fontWeight: 600,
              cursor: "pointer",
            }}
          >
            Sign in to save
          </button>
        ) : readOnly ? (
          <span style={{ fontSize: 13, fontWeight: 600, color: "#64748b" }}>
            View only
          </span>
        ) : (
          <button
            type="button"
            disabled={saving}
            onClick={handleSave}
            style={{
              padding: "9px 16px",
              border: "none",
              borderRadius: 8,
              background: "#059669",
              color: "#fff",
              fontWeight: 600,
              cursor: saving ? "default" : "pointer",
              opacity: saving ? 0.7 : 1,
            }}
          >
            {saving ? "Saving…" : "Save"}
          </button>
        )}
      </footer>
    </div>
  );
};

export default ThetaSharePage;
