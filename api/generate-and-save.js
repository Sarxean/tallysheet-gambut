// pages/api/generate-and-save.js

import { IncomingForm } from "formidable";
import htmlToDocx from "html-to-docx";
import { createClient } from "@supabase/supabase-js";

// Nonaktifkan bodyParser bawaan Next.js
export const config = { api: { bodyParser: false } };

// Fungsi set header CORS
function setCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

// Sanitizer HTML yang kuat untuk menghapus atribut/tagnames tidak valid (mis. @w, w:xxx, x:xxx) dan karakter terlarang.
function sanitizeHtmlContent(input) {
  let str = String(input || "");

  // Hapus XML declaration jika ada
  str = str.replace(/<\?xml[^>]*\?>/gi, "");

  // Hapus atribut yang diawali '@' (contoh: @w, @x, dst)
  str = str.replace(/\s+@[^\s=>/]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?/g, "");

  // Hapus atribut ber-namespace (contoh: w:val="..", x:width="..", termasuk xmlns:)
  str = str.replace(/\s+(?:xmlns(?::\w+)?|[A-Za-z_][\w-]*:[\w-]+)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/g, "");

  // Hapus tag dengan nama ber-namespace atau diawali '@' (contoh: <w:tbl>..</w:tbl>, <@w ...>)
  str = str.replace(/<\/?\s*(?:@[A-Za-z_][\w-]*|[A-Za-z_][\w-]*:[\w-]+)[^>]*>/g, "");

  // Bersihkan token @w:xxx di dalam nilai atribut/style/text agar aman
  str = str.replace(/@\w+:[^;"'>\s]+;?/g, "");

  // Hapus karakter kontrol tak valid untuk XML/Word
  str = str.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "");

  return str;
}

export default async function handler(req, res) {
  setCors(res);

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ success: false, message: "Method Not Allowed" });
  }

  try {
    const {
      SUPABASE_URL,
      SUPABASE_SERVICE_ROLE_KEY,
      SUPABASE_BUCKET_DOC = "tallysheet-files",
      SIGNED_URL_EXPIRE = "3600",
    } = process.env;

    if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ success: false, message: "Missing Supabase config" });
    }

    const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
    const signedExpire = parseInt(SIGNED_URL_EXPIRE, 10) || 3600;

    let html = "";
    let fileNameField = "";
    let userIdField = "";
    let folderIdField = "";
    let tallyIdField = "";

    // Parsing body sesuai Content-Type
    if (req.headers["content-type"]?.includes("application/json")) {
      const body = await new Promise((resolve, reject) => {
        let raw = "";
        req.on("data", (chunk) => (raw += chunk));
        req.on("end", () => {
          try {
            resolve(JSON.parse(raw));
          } catch (e) {
            reject(e);
          }
        });
      });

      html = body?.html || "";
      fileNameField = body?.file_name;
      userIdField = body?.user_id;
      folderIdField = body?.folder_id;
      tallyIdField = body?.tallysheet_id;
    } else if (req.headers["content-type"]?.includes("multipart/form-data")) {
      const form = new IncomingForm({ multiples: false, keepExtensions: true });
      const { fields } = await new Promise((resolve, reject) => {
        form.parse(req, (err, flds, fls) => {
          if (err) reject(err);
          else resolve({ fields: flds, files: fls });
        });
      });

      const getField = (field) =>
        Array.isArray(fields?.[field]) ? fields[field][0] : fields?.[field];

      html = getField("html");
      fileNameField = getField("file_name");
      userIdField = getField("user_id");
      folderIdField = getField("folder_id");
      tallyIdField = getField("tallysheet_id");
    }

    if (!html) {
      return res.status(400).json({ success: false, message: "Missing 'html' field" });
    }

    // SANITASI HTML sebelum convert -> cegah Invalid XML name (@w, w:xxx, dll)
    html = sanitizeHtmlContent(html);

    // Validasi dan sanitasi NAMA FILE
    let safeFileName;
    try {
      let rawName = Array.isArray(fileNameField) ? fileNameField[0] : fileNameField;
      rawName = String(rawName ?? "").trim();

      // Ganti karakter terlarang (Windows reserved) & kontrol
      safeFileName = rawName
        .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
        .replace(/^\.+$/, "_") // kalau hanya titik
        .replace(/\s+/g, " ") // normalisasi spasi
        .trim();

      if (!safeFileName) safeFileName = `tallysheet_${Date.now()}`;
      if (!safeFileName.toLowerCase().endsWith(".docx")) safeFileName += ".docx";
    } catch (_) {
      safeFileName = `tallysheet_${Date.now()}.docx`;
    }

    const userId = String(userIdField || "anonymous").replace(/[\\/]/g, "_");
    const storagePath = `${userId}/${safeFileName}`;

    // Convert HTML ke DOCX (Buffer → Uint8Array)
    const docxBuffer = await htmlToDocx(html, null, {
      table: { row: { cantSplit: true } },
      footer: true,
      pageNumber: true,
    });

    const uint8 = docxBuffer instanceof Uint8Array ? docxBuffer : new Uint8Array(docxBuffer);

    // Upload file ke Supabase Storage (UPSERT)
    const { error: uploadErr } = await supabase.storage
      .from(SUPABASE_BUCKET_DOC)
      .upload(storagePath, uint8, {
        upsert: true,
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

    if (uploadErr) {
      return res
        .status(500)
        .json({ success: false, message: "Upload failed", error: uploadErr.message });
    }

    const metadata = {
      tallysheet_id: tallyIdField || null,
      uploaded_by: userId,
      uploaded_at: new Date().toISOString(),
    };

    // Cek apakah file sudah ada di tabel "files"
    const { data: existing, error: existingErr } = await supabase
      .from("files")
      .select("id")
      .eq("path", storagePath)
      .limit(1);

    if (existingErr) {
      console.error("Error checking existing file:", existingErr);
    }

    if (existing && existing.length > 0) {
      await supabase
        .from("files")
        .update({
          name: safeFileName,
          folder_id: folderIdField || null,
          metadata,
          updated_at: new Date().toISOString(),
        })
        .eq("id", existing[0].id);
    } else {
      await supabase.from("files").insert([
        {
          user_id: userId,
          name: safeFileName,
          path: storagePath,
          folder_id: folderIdField || null,
          metadata,
          created_at: new Date().toISOString(),
          updated_at: new Date().toISOString(),
        },
      ]);
    }

    // Buat Signed URL
    const { data: signedData, error: signedErr } = await supabase.storage
      .from(SUPABASE_BUCKET_DOC)
      .createSignedUrl(storagePath, signedExpire);

    if (signedErr) {
      console.error("Error creating signed URL:", signedErr);
    }

    return res.status(200).json({
      success: true,
      path: storagePath,
      file_name: safeFileName,
      signed_url: signedData?.signedURL || signedData?.signed_url || null,
    });
  } catch (err) {
    console.error("Server error:", err);
    setCors(res);
    return res.status(500).json({ success: false, message: err.message || "Server error" });
  }
}