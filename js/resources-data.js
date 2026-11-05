/* ═══════════════════════════════════════════════════════════════════════
   resources-data.js — DANH SÁCH HỌC LIỆU SỐ  (file bạn sẽ chỉnh sửa)

   Mỗi học liệu là MỘT khối { ... } trong mảng CHEMICAL_RESOURCES bên dưới.
   Chỉ cần COPY một khối mẫu, dán xuống cuối mảng (nhớ có dấu phẩy) và sửa.

   ── CÁC TRƯỜNG ──────────────────────────────────────────────────────────
   ResourceID    : số thứ tự không trùng nhau (1, 2, 3, ...)
   ResourceTypeID: 1 = Video tuyên truyền
                   2 = Infographic (hiển thị ảnh trực tiếp trên web)
                   3 = Tài liệu hướng dẫn (PDF)
   Title         : tiêu đề hiển thị
   Description   : mô tả ngắn (không bắt buộc)
   FileURL       : • Video / PDF  → dán LINK GOOGLE DRIVE (hoặc link YouTube)
                   • Infographic  → đường dẫn ảnh, ví dụ
                                    "assets/infographics/cach-doc-nhan-ghs.png"
                                    (copy ảnh vào thư mục assets/infographics/)
   ChemicalID    : (không bắt buộc) mã hóa chất liên quan, ví dụ "C003".
                   Nếu có, học liệu sẽ hiện thêm trong trang hồ sơ hóa chất
                   đó. Học liệu chung thì để null.

   ── CÁCH LẤY LINK GOOGLE DRIVE ──────────────────────────────────────────
   1. Tải video/PDF lên Drive.
   2. Chuột phải file → Chia sẻ → "Quyền truy cập chung" chọn
      "Bất kỳ ai có đường liên kết" (quyền Người xem).
   3. Chọn "Sao chép đường liên kết" rồi dán vào FileURL.
   (Nếu không bật chia sẻ công khai, người xem sẽ không mở được.)

   ── MẪU (bỏ dấu // ở đầu mỗi dòng để dùng) ─────────────────────────────
   // { ResourceID: 1, ResourceTypeID: 1, ChemicalID: null,
   //   Title: "Video: Không trộn nước tẩy với axit",
   //   Description: "Video ngắn về nguy cơ khi trộn hóa chất tẩy rửa.",
   //   FileURL: "https://drive.google.com/file/d/XXXXXXXX/view?usp=sharing" },
   //
   // { ResourceID: 2, ResourceTypeID: 2, ChemicalID: null,
   //   Title: "Infographic: Cách đọc nhãn GHS",
   //   Description: "9 biểu tượng nguy hiểm GHS và ý nghĩa.",
   //   FileURL: "assets/infographics/cach-doc-nhan-ghs.png" },
   //
   // { ResourceID: 3, ResourceTypeID: 3, ChemicalID: "C003",
   //   Title: "Sổ tay an toàn hóa chất gia dụng",
   //   Description: "Tài liệu PDF hướng dẫn sử dụng và xử lý sự cố.",
   //   FileURL: "https://drive.google.com/file/d/YYYYYYYY/view?usp=sharing" },
   ═══════════════════════════════════════════════════════════════════════ */

   

const CHEMICAL_RESOURCES = [
  // ← dán các học liệu của bạn vào đâynst CHEMICAL_RESOURCES = [
  { ResourceID: 1, ResourceTypeID: 1, ChemicalID: null,
    Title: "Video: Không trộn nước tẩy với axit",
    Description: "Video ngắn về nguy cơ khi trộn hóa chất tẩy rửa.",
    FileURL: "https://drive.google.com/drive/folders/114Bp5FY2mFaOk-eb-XdlXUZE0efqIJF7?usp=sharing" },
   
  { ResourceID: 2, ResourceTypeID: 2, ChemicalID: null,
     Title: "Infographic 1: Đọc nhãn trước khi dùng",
     Description: "9 biểu tượng nguy hiểm GHS và ý nghĩa.",
     FileURL: "assets/infographics/photo1.jpg" },
     
  { ResourceID: 3, ResourceTypeID: 2, ChemicalID: null,
    Title: "Infographic 2: Lưu ý thành phần khi chọn sản phẩm",
    Description: "Mô tả ngắn về nội dung ảnh.",
    FileURL: "assets/infographics/photo2.jpg" },

  { ResourceID: 4, ResourceTypeID: 2, ChemicalID: null,
    Title: "Infographic 3: Không pha trộn chất tẩy rửa",
    Description: "Mô tả ngắn về nội dung ảnh.",
    FileURL: "assets/infographics/photo3.jpg" },

  { ResourceID: 5, ResourceTypeID: 2, ChemicalID: null,
    Title: "Infographic 4: Dùng đúng cách bình xịt diệt côn trùng",
    Description: "Mô tả ngắn về nội dung ảnh.",
    FileURL: "assets/infographics/photo4.jpg" },

  { ResourceID: 6, ResourceTypeID: 2, ChemicalID: null,
    Title: "Infographic 5: Dùng đúng cách bình xịt gia trùng",
    Description: "Mô tả ngắn về nội dung ảnh.",
    FileURL: "assets/infographics/photo41.jpg" },

  { ResourceID: 7, ResourceTypeID: 2, ChemicalID: null,
    Title: "Infographic 6: Bảo vệ nguồn nước",
    Description: "Mô tả ngắn về nội dung ảnh.",
    FileURL: "assets/infographics/photo5.jpg" },

   { ResourceID: 8, ResourceTypeID: 3, ChemicalID: "C003",
     Title: "Sổ tay an toàn hóa chất gia dụng",
      Description: "Tài liệu PDF hướng dẫn sử dụng và xử lý sự cố.",
     FileURL: "https://drive.google.com/drive/u/0/folders/1id5tFeyeymJfMrTOAsRTnZJovjG0EMqH" },
];

/* ═══════════════════════════════════════════════════════════════════════
   HÀM HỖ TRỢ (không cần chỉnh sửa)
   ═══════════════════════════════════════════════════════════════════════ */

/* Lấy ID file từ các dạng link Google Drive */
function driveFileId(url) {
  if (!url) return null;
  const m = String(url).match(/drive\.google\.com\/(?:file\/d\/|open\?id=|uc\?(?:[^#]*&)?id=)([\w-]{10,})/);
  return m ? m[1] : null;
}

/* Lấy ID video từ link YouTube */
function youtubeId(url) {
  if (!url) return null;
  const m = String(url).match(/(?:youtu\.be\/|youtube\.com\/(?:watch\?(?:[^#]*&)?v=|embed\/|shorts\/))([\w-]{11})/);
  return m ? m[1] : null;
}

/* Link mở học liệu ở tab mới */
function resourceOpenUrl(r) {
  const id = driveFileId(r.FileURL);
  if (id) return `https://drive.google.com/file/d/${id}/view`;
  return r.FileURL;
}

/* Link nhúng (iframe) để xem ngay trên web; null nếu không nhúng được */
function resourceEmbedUrl(r) {
  const d = driveFileId(r.FileURL);
  if (d) return `https://drive.google.com/file/d/${d}/preview`;
  const y = youtubeId(r.FileURL);
  if (y) return `https://www.youtube.com/embed/${y}`;
  return null;
}

/* Ảnh đại diện của thẻ học liệu; null nếu không có */
function resourceThumbUrl(r) {
  if (r.ResourceTypeID === 2) return r.FileURL;          // infographic: chính là ảnh
  const d = driveFileId(r.FileURL);
  if (d) return `https://drive.google.com/thumbnail?id=${d}&sz=w640`;
  const y = youtubeId(r.FileURL);
  if (y) return `https://img.youtube.com/vi/${y}/hqdefault.jpg`;
  return null;
}
