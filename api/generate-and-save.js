// api/generate-and-save.js
import { IncomingForm } from "formidable";
import fs from "fs";
import htmlToDocx from "html-to-docx";
import { createClient } from "@supabase/supabase-js";
import fetch from "node-fetch";

// disable body parser for Vercel
export const config = { api: { bodyParser: false } };

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const BUCKET_DOC = process.env.SUPABASE_BUCKET_DOC || "tallysheet-files";
const BUCKET_IMG = process.env.SUPABASE_BUCKET_IMG || "fotos";
const SIGN_EXPIRE = parseInt(process.env.SIGNED_URL_EXPIRE || "3600", 10);

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  throw new Error("Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY in env");
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);

function parseForm(req) {
  const form = new IncomingForm();
  form.keepExtensions = true;
  return new Promise((resolve, reject) => {
    form.parse(req, (err, fields, files) => {
      if (err) reject(err);
      else resolve({ fields, files });
    });
  });
}

async function fetchAsDataURL(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`Failed fetch image ${url}: ${r.status}`);
  const buff = await r.arrayBuffer();
  const b64 = Buffer.from(buff).toString("base64");
  const contentType = r.headers.get("content-type") || "image/png";
  return `data:${contentType};base64,${b64}`;
}

export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      res.status(405).json({ message: "Method not allowed" });
      return;
    }

    const { fields } = await parseForm(req);
    const html = fields.html;
    if (!html) return res.status(400).json({ message: "Missing html field" });

    const fileNameRaw = fields.file_name || `tallysheet_${Date.now()}.docx`;
    const safeFileName = fileNameRaw.replace(/[\\/]+/g, "_");
    const userId = fields.user_id || "anonymous";
    const tallysheetId = fields.tallysheet_id || null;
    const folderId = fields.folder_id || null;

    // STEP: replace non-data img src with data URLs by fetching them
    // This tries to fetch each <img src="..."> that isn't already data:
    // - If src is accessible (public or signed) it will be converted to data URL
    // - If not accessible (fetch fails), but we have tallysheetId, we will try to read image path from DB (tallysheet table)
    // Simple parsing by regex to find img src occurrences:
    let processedHtml = html;
    const imgSrcRegex = /<img[^>]+src=["']([^"']+)["'][^>]*>/gi;
    const srcs = [];
    let m;
    while ((m = imgSrcRegex.exec(html)) !== null) {
      const src = m[1];
      if (src && !src.startsWith("data:")) srcs.push(src);
    }

    // try to fetch each src and replace with data URL
    for (const src of srcs) {
      try {
        const dataUrl = await fetchAsDataURL(src);
        // replace all occurrences of this src with dataUrl (be careful with characters)
        processedHtml = processedHtml.split(src).join(dataUrl);
      } catch (err) {
        console.warn("Could not fetch image from src:", src, err.message);
        // If fetch failed, and tallysheet_id provided, try to map by filename/key
        // We'll attempt to find a matching column in tallysheet table whose filename equals end of src.
        // Otherwise leave src as-is (html-to-docx may attempt to fetch it).
        if (tallysheetId) {
          try {
            const { data: tsRow, error: tsError } = await supabase
              .from("tallysheet")
              .select("*")
              .eq("id", tallysheetId)
              .single();
            if (tsError) throw tsError;
            // Try to find matching image column by checking values in tsRow that equal or endWith src basename
            const basename = src.split("/").pop();
            let matchedPath = null;
            for (const [col, val] of Object.entries(tsRow)) {
              if (!val) continue;
              if (typeof val === "string" && (val === src || val.endsWith(basename) || val.includes(basename))) {
                matchedPath = val; // assume this is storage path
                break;
              }
            }
            if (matchedPath) {
              // download from supabase storage
              // matchedPath might be full path like "userId/filename.jpg" or public URL
              let downloadBuffer = null;
              // If matchedPath looks like full URL, try fetching
              if (matchedPath.startsWith("http")) {
                try {
                  const r2 = await fetch(matchedPath);
                  const buff2 = await r2.arrayBuffer();
                  downloadBuffer = Buffer.from(buff2);
                } catch (e2) {
                  console.warn("failed to fetch matchedPath URL", e2.message);
                }
              } else {
                // assume it's storage path in bucket BUCKET_IMG
                const { data: downloaded, error: dlErr } = await supabase.storage
                  .from(BUCKET_IMG)
                  .download(matchedPath);
                if (dlErr) {
                  console.warn("Supabase download failed for", matchedPath, dlErr.message);
                } else {
                  // downloaded is a ReadableStream in browser, but in Node environment supabase-js returns Buffer-like
                  const arrayBuffer = await downloaded.arrayBuffer();
                  downloadBuffer = Buffer.from(arrayBuffer);
                }
              }
              if (downloadBuffer) {
                const contentType = "image/png"; // fallback
                const b64 = downloadBuffer.toString("base64");
                const dataUrl = `data:${contentType};base64,${b64}`;
                processedHtml = processedHtml.split(src).join(dataUrl);
              }
            }
          } catch (e) {
            console.warn("Mapping via tallysheet failed:", e.message);
          }
        }
      }
    }

    // Convert the processed HTML to DOCX buffer
    const fileBuffer = await htmlToDocx(processedHtml, null, {
      // options if needed
    });

    // Upload docx to supabase bucket BUCKET_DOC
    const storagePath = `${userId}/${safeFileName}`;
    // convert Buffer to Uint8Array for supabase client
    const uint8 = Uint8Array.from(fileBuffer);

    const { error: uploadErr } = await supabase.storage
      .from(BUCKET_DOC)
      .upload(storagePath, uint8, {
        upsert: true,
        contentType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

    if (uploadErr) {
      console.error("Upload error:", uploadErr);
      return res.status(500).json({ message: "Upload failed", error: uploadErr });
    }

    // Insert a record into files (or update if exists — we'll upsert by path)
    // Simple approach: try insert; if conflict on path, update
    // Ensure table 'files' exists with unique constraint on path or implement simple upsert via select->update/insert
    const metadata = {
      tallysheet_id: tallysheetId,
      uploaded_by: userId,
      uploaded_at: new Date().toISOString(),
    };

    // Try to find existing file row
    const { data: existingRows } = await supabase
      .from("files")
      .select("*")
      .eq("path", storagePath)
      .limit(1);

    if (existingRows && existingRows.length > 0) {
      const existing = existingRows[0];
      const { error: updErr } = await supabase
        .from("files")
        .update({ name: safeFileName, folder_id: folderId || null, metadata })
        .eq("id", existing.id);
      if (updErr) console.warn("files update warning:", updErr);
    } else {
      const { error: insErr } = await supabase.from("files").insert([
        {
          user_id: userId,
          name: safeFileName,
          path: storagePath,
          folder_id: folderId || null,
          metadata,
        },
      ]);
      if (insErr) console.warn("files insert warning:", insErr);
    }

    // create signed URL for preview
    const { data: signedData, error: signedErr } = await supabase.storage
      .from(BUCKET_DOC)
      .createSignedUrl(storagePath, SIGN_EXPIRE);
    if (signedErr) {
      console.warn("Signed url warning:", signedErr);
    }

    return res.json({
      success: true,
      path: storagePath,
      signed_url: signedData?.signedURL || signedData?.signed_url || null,
    });
  } catch (err) {
    console.error("Server error:", err);
    return res.status(500).json({ message: "Server error", error: err.message });
  }
}