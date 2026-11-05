/* ═══════════════════════════════════════════════════════════════════════
   data.js — CƠ SỞ DỮ LIỆU HÓA CHẤT TRONG SINH HOẠT  (tự sinh từ file Excel)
   Nguồn: Co_so_du_lieu_108_san_pham_199_hoa_chat_HPL_ra_soat_03-11-2026.xlsx
   Phạm vi: 108 sản phẩm · 199 hóa chất (143 gốc + 56 mới) · 573 liên kết
   Phân mức HPL của 199 chất: HPL-3: 36 · HPL-2: 69 · HPL-1: 67 · Chưa phân mức: 27

   ⚠ QUAN TRỌNG (theo README của bộ dữ liệu):
   - Kết quả bài báo gốc chỉ tính trên 143 chất (31 HPL-3 · 54 HPL-2 · 58 HPL-1).
     KHÔNG thay số liệu bài báo bằng bộ 199 chất.
   - Hóa chất mới (C144–C199) thiếu nguồn GHS/SDS vẫn để "CHƯA PHÂN MỨC"
     (HPLLevelID = 0) — thiếu dữ liệu KHÔNG đồng nghĩa HPL-1/an toàn.
   - Hóa chất mới chỉ xác minh sự có mặt trong sản phẩm; nhiều trường GHS /
     độc tính sinh thái còn "chưa rà soát" (để trống).
   - HPL sản phẩm là mức SUY RA TỪ THÀNH PHẦN, không phải phân loại GHS của
     hỗn hợp; chưa xét nồng độ, liều, đường và thời gian phơi nhiễm.
     "Tối thiểu HPL-x" = còn chất chưa phân mức; "Chưa xác định" = chưa đủ
     thành phần đã xác minh.
   - "NF" = chưa tìm thấy dữ liệu định lượng phù hợp, không đồng nghĩa không độc.
   ═══════════════════════════════════════════════════════════════════════ */

/* ═══════════════════════════════════════════════════════════════════════
   1. CATEGORY — Nhóm sản phẩm sinh hoạt (5 nhóm, trích từ sheet PRODUCTS)
   ═══════════════════════════════════════════════════════════════════════ */
const CATEGORIES = [
  {
    "CategoryID": 1,
    "CategoryName": "Chất tẩy rửa và vệ sinh gia dụng",
    "Description": "Nước rửa chén, nước giặt, nước lau sàn, nước tẩy nhà vệ sinh, xịt khử mùi...",
    "icon": "🧴"
  },
  {
    "CategoryID": 2,
    "CategoryName": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
    "Description": "Dầu gội, sữa tắm, kem dưỡng da, sữa rửa mặt, nước hoa, mỹ phẩm...",
    "icon": "🧼"
  },
  {
    "CategoryID": 3,
    "CategoryName": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
    "Description": "Thuốc trừ sâu, thuốc trừ cỏ, bình xịt côn trùng...",
    "icon": "🌿"
  },
  {
    "CategoryID": 4,
    "CategoryName": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
    "Description": "Phụ gia thực phẩm, màu thực phẩm, hương liệu, kem/bột làm bánh, đồ hộp...",
    "icon": "🍽️"
  },
  {
    "CategoryID": 5,
    "CategoryName": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
    "Description": "Sơn, keo dán, dầu phanh, bình xịt bôi trơn đa năng...",
    "icon": "⛽"
  }
];
function getCategoryById(id) { return CATEGORIES.find(c => c.CategoryID === parseInt(id)) || null; }
function getCategoryByName(name) { return CATEGORIES.find(c => c.CategoryName === name) || null; }


/* ═══════════════════════════════════════════════════════════════════════
   2. HPL_LEVEL — Hệ thống phân mức cảnh báo (nguyên tắc tổng hợp, theo
   đúng sheet HPL_RULES của bộ dữ liệu gốc — "chọn mức nguy hại cao nhất
   có đủ bằng chứng; không cộng điểm cơ học giữa các tiêu chí").
   ═══════════════════════════════════════════════════════════════════════ */
const HPL_LEVELS = [
  {
    HPLLevelID: 0, HPLCode: "CHƯA PHÂN MỨC", LevelName: "Chưa phân mức HPL",
    ColorCode: "#6b7280", cssClass: "level-unrated", icon: "⚪",
    Description: "Chưa đủ bằng chứng GHS/SDS để xác định mức HPL. Thiếu dữ " +
      "liệu KHÔNG được xem là bằng chứng an toàn và không được tự gán HPL-1.",
  },
  {
    HPLLevelID: 3, HPLCode: "HPL-3", LevelName: "Ưu tiên cảnh báo cao",
    ColorCode: "#dc2626", cssClass: "level-danger", icon: "🔴",
    Description: "Có ít nhất một căn cứ đủ độ tin cậy ở mức cao: độc tính " +
      "cấp GHS mức nghiêm trọng, gây ung thư/đột biến/độc tính sinh sản, " +
      "độc tính đặc hiệu cơ quan đích nghiêm trọng, phân loại nguy hại " +
      "thủy sinh Aquatic Acute 1 (H400) hoặc Aquatic Chronic 1 (H410), " +
      "độc tính cấp thủy sinh LC₅₀/EC₅₀ ≤ 1 mg/L, hoặc độc tính mạn " +
      "NOEC/EC₁₀ ≤ 0,1 mg/L (chất không phân hủy nhanh) / ≤ 0,01 mg/L " +
      "(chất phân hủy nhanh).",
  },
  {
    HPLLevelID: 2, HPLCode: "HPL-2", LevelName: "Ưu tiên cảnh báo trung bình",
    ColorCode: "#d97706", cssClass: "level-warn", icon: "🟠",
    Description: "Không đạt căn cứ HPL-3 nhưng có ít nhất một căn cứ HPL-2 " +
      "đủ độ tin cậy: độc tính cấp GHS mức thấp hơn, ăn mòn/kích ứng, " +
      "mẫn cảm, độc tính cơ quan đích cần lưu ý, phân loại Aquatic " +
      "Chronic 2–4 (H411–H413), độc tính cấp thủy sinh > 1–100 mg/L, hoặc " +
      "độc tính mạn > 0,1–1 mg/L (chất không phân hủy nhanh).",
  },
  {
    HPLLevelID: 1, HPLCode: "HPL-1", LevelName: "Ưu tiên cảnh báo thấp",
    ColorCode: "#16a34a", cssClass: "level-safe", icon: "🟢",
    Description: "Không ghi nhận căn cứ HPL-2/HPL-3 từ dữ liệu được lựa " +
      "chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo THẤP trên dữ liệu " +
      "hiện có — KHÔNG đồng nghĩa hóa chất an toàn (có thể do thiếu dữ " +
      "liệu định lượng chuẩn hóa, không phải do đã được chứng minh an toàn).",
  },
];
function getHPLLevel(id) { return HPL_LEVELS.find(h => h.HPLLevelID === parseInt(id)) || null; }

/* Phạm vi/giới hạn của thang HPL (theo sheet HPL_RULES, dòng "Phạm vi") */
const HPL_SCOPE_NOTE = "HPL là thang ưu tiên cảnh báo nội bộ phục vụ truyền " +
  "thông/tra cứu, KHÔNG phải phân loại chính thức của GHS và không phải " +
  "đánh giá rủi ro có xét đến mức độ phơi nhiễm thực tế. Nguy hại vật lý " +
  "(cháy, nổ, oxy hóa, khí chịu áp suất...) được cảnh báo riêng theo " +
  "GHS/SDS, không quy đổi trực tiếp vào thang HPL này.";


/* ═══════════════════════════════════════════════════════════════════════
   3. GHS_PICTOGRAMS — Bảng tra icon cho 9 biểu tượng nguy hiểm GHS chuẩn
   ═══════════════════════════════════════════════════════════════════════ */
const GHS_PICTOGRAMS = {
  GHS01: { name: "Chất nổ",                       icon: "💥" },
  GHS02: { name: "Chất dễ cháy",                   icon: "🔥" },
  GHS03: { name: "Chất oxy hóa",                   icon: "⭕" },
  GHS04: { name: "Khí nén",                        icon: "🧯" },
  GHS05: { name: "Ăn mòn",                         icon: "🧪" },
  GHS06: { name: "Độc cấp tính",                   icon: "☠️" },
  GHS07: { name: "Kích ứng / Có hại",              icon: "⚠️" },
  GHS08: { name: "Nguy hại sức khỏe nghiêm trọng", icon: "🫁" },
  GHS09: { name: "Nguy hại môi trường",            icon: "🌊" },
};


/* ═══════════════════════════════════════════════════════════════════════
   4. RESOURCE_TYPE / CHEMICAL_RESOURCE — Học liệu số (Mục 2.4)
   ⚠️ Bộ dữ liệu bài báo không bao gồm học liệu số — để trống, cấu trúc
   đã sẵn sàng để bổ sung video/infographic/PDF khi có.
   ═══════════════════════════════════════════════════════════════════════ */
const RESOURCE_TYPES = [
  { TypeID: 1, TypeName: "Video hướng dẫn",       icon: "🎬", Description: "Video giới thiệu sản phẩm, thành phần, nguy cơ, hướng dẫn sử dụng an toàn." },
  { TypeID: 2, TypeName: "Infographic",              icon: "🖼️", Description: "Infographic biểu tượng GHS, nguyên tắc an toàn, cách đọc nhãn." },
  { TypeID: 3, TypeName: "Sổ tay an toàn hoá chất gia dụng", icon: "📄", Description: "Sổ tay tra cứu hóa chất nguy cơ, hướng dẫn phòng tránh và xử lý rủi ro." },
];
function getResourceType(id) { return RESOURCE_TYPES.find(r => r.TypeID === parseInt(id)) || null; }
/* CHEMICAL_RESOURCES nằm trong js/resources-data.js */
function getResourcesByChemicalId(chemicalId) {
  return CHEMICAL_RESOURCES
    .filter(r => r.ChemicalID === chemicalId)
    .map(r => ({ ...r, resourceType: getResourceType(r.ResourceTypeID) }));
}



/* ═══════════════════════════════════════════════════════════════════════
   5. PRODUCT — 108 sản phẩm (kèm HPL suy ra từ thành phần)
   ═══════════════════════════════════════════════════════════════════════ */
const PRODUCTS = [
 {
  "ProductID": "P001",
  "CategoryID": 1,
  "ProductName": "Bình xịt khử mùi giày dép Ion Bạc",
  "Manufacturer": "Công ty TNHH Công nghệ hóa chất Ngự Thượng Quảng Châu",
  "SourceURL": "https://tatoshopvn.com/binh-xit-khu-mui-giay-dep-nano-bac-260ml?srsltid=AU7gw4UWiMBgLghyMx_I-CLO6QeER2B2NQiQqyRtgElO5XzyUUMXyAsP",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 1,
  "NumUnrated": 1,
  "AssessmentStatus": "CẦN XÁC MINH THÀNH PHẦN",
  "DominantChemicals": "CHƯA XÁC MINH (nano bạc trong tên thương mại)",
  "ScopeNote": "Chỉ có tên thương mại ghi nano bạc; chưa có nhãn/SDS và đặc trưng vật liệu nano. Không dùng HPL của bạc khối để gán HPL sản phẩm."
 },
 {
  "ProductID": "P002",
  "CategoryID": 1,
  "ProductName": "Bình xịt khử mùi phòng Taman",
  "Manufacturer": "Công ty TNHH / Công nghệ Hữu cơ Taman",
  "SourceURL": "https://shopee.vn/X%E1%BB%8Bt-Kh%E1%BB%AD-M%C3%B9i-Ph%C3%B2ng-X%E1%BB%8Bt-Ph%C3%B2ng-Th%C6%A1m-TAMAN-Cao-C%E1%BA%A5p-H%C6%B0%C6%A1ng-Sen-Nano-B%E1%BA%A1c-Kh%C3%A1ng-Khu%E1%BA%A9n-Hi%E1%BB%87u-Qu%E1%BA%A3-An-To%C3%A0n-i.1203778738.41418780752",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P003",
  "CategoryID": 1,
  "ProductName": "Bình xịt thơm tủ quần áo VELVET MIST ABURA",
  "Manufacturer": "Công ty TNHH Vina Tươi",
  "SourceURL": "https://abura.vn/san-pham/xit-khu-mui-quan-ao-velvet-mist-abura",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P004",
  "CategoryID": 1,
  "ProductName": "Chai nước tẩy rửa bồn cầu Okey Pink nhập khẩu từ Thái Lan",
  "Manufacturer": "White House",
  "SourceURL": "https://truonganmart.vn/nuoc-tay-toilet-okay-pink-960ml-thai-lan-chinh-hang-12-1/17011009000894046",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C062 – Hydrochloric acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P005",
  "CategoryID": 1,
  "ProductName": "Chai tẩy rửa inox Denkmit Edelstahlreiniger 4in1",
  "Manufacturer": "Đức",
  "SourceURL": "https://minhhouseware.com.vn/san-pham/chai-tay-rua-inox-denkmit-3in1",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C116 – Sodium pyrithione",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P006",
  "CategoryID": 1,
  "ProductName": "Chai xịt lau chùi và khử mùi bếp điện Thơm, giúp vệ sinh bề mặt bếp nhanh chóng",
  "Manufacturer": null,
  "SourceURL": null,
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P007",
  "CategoryID": 1,
  "ProductName": "Chai xịt tẩy cặn canxi Antikal Kalkreiniger Fresh",
  "Manufacturer": "P&G (Procter & Gamble)",
  "SourceURL": "https://sieuthidanke.vn/san-pham/chai-xit-antikal-ve-sinh-va-chong-an-mon-inox/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "CẦN XÁC MINH THÀNH PHẦN",
  "DominantChemicals": "C036 – Citric acid; C166 – Formic acid; C159 – Deceth-8",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp. CẦN XÁC MINH NGUỒN THÀNH PHẦN: C036: SDS Antikal Fresh mit Febreze 2022; cần đối chiếu mã lô/phiên bản với chai bán tại Việt Nam.; C166: SDS Antikal Fresh mit Febreze 2022; cần đối chiếu mã lô/phiên bản với chai bán tại Việt Nam.; C159: SDS Antikal Fresh mit Febreze 2022; cần đối chiếu mã lô/phiên bản với chai bán tại Việt Nam."
 },
 {
  "ProductID": "P008",
  "CategoryID": 1,
  "ProductName": "Chai xịt vệ sinh lò nướng và vỉ nướng Reine Welle.",
  "Manufacturer": "Reine Welle (Đức)",
  "SourceURL": null,
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P009",
  "CategoryID": 1,
  "ProductName": "Nước lau bếp Sunlight Bột Tuyết Powerfoam",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": "https://www.unilever.com.vn",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C089 – Phenoxyethanol; C104 – Sodium carbonate; C109 – Sodium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P010",
  "CategoryID": 1,
  "ProductName": "Nước lau sàn Sunlight Bioshield hương Chanh Yuzu & Sả",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": null,
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol; C109 – Sodium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P011",
  "CategoryID": 1,
  "ProductName": "Nước tẩy bồn cầu Vim Trắng sáng Hương Chanh Sả",
  "Manufacturer": "Công ty Cổ phần Bột giặt Lix - LIXCO",
  "SourceURL": "https://www.lixco.com",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C028 – Cetrimonium chloride",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P012",
  "CategoryID": 1,
  "ProductName": "Nước tẩy bồn cầu và nhà tắm Vim",
  "Manufacturer": "Unilever Việt Nam",
  "SourceURL": "https://www.unilever.com.vn/brands/home-care/vim/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C110 – Sodium hypochlorite",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P013",
  "CategoryID": 1,
  "ProductName": "Nước rửa chén Gift hương chanh",
  "Manufacturer": "Công ty TNHH Earth Corporation Việt Nam",
  "SourceURL": "https://www.myhao.vn/san-pham/nuoc-rua-chen-my-hao",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 8,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C079 – Methylchloroisothiazolinone; C080 – Methylisothiazolinone",
  "ScopeNote": "Có chất chưa phân mức; đã có bằng chứng HPL-3 ở ít nhất một thành phần."
 },
 {
  "ProductID": "P014",
  "CategoryID": 1,
  "ProductName": "Nước rửa chén Lix siêu đậm đặc hương trà xanh",
  "Manufacturer": "Công ty Cổ phần Bột giặt Lix",
  "SourceURL": "https://lixco.com/san-pham/lix/rua-chen/sieu-dam-dac/nuoc-rua-chen-lix-sieu-dam-dac-huong-tra-xanh-172.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 7,
  "NumUnrated": 2,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C151 – Benzisothiazolinone",
  "ScopeNote": "Có chất chưa phân mức; đã có bằng chứng HPL-3 ở ít nhất một thành phần."
 },
 {
  "ProductID": "P015",
  "CategoryID": 1,
  "ProductName": "Nước thông cống Blue",
  "Manufacturer": "Blue Chemical Co., Ltd. (Thái Lan)",
  "SourceURL": "https://mmpro.vn/product/nuoc-thong-cong-blue-1l-479962-24799621.html?srsltid=AfmBOopCStUbPGxTVlSwhKe5jPHNcZV3WJ6Wy3s5SnUJGy11QAo2vAWg",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P016",
  "CategoryID": 1,
  "ProductName": "Nước lau sàn Power100 hương tinh dầu chanh sả.",
  "Manufacturer": "Power100",
  "SourceURL": "https://www.calofic.com.vn/product/nuoc-lau-nha-power100-huong-chanh-sa/index.html",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P017",
  "CategoryID": 1,
  "ProductName": "Chai xịt mạt bụi vải KA (KA Anti-Dust Mite Fabric Spray)",
  "Manufacturer": "OUJI / KA (thương hiệu KA)",
  "SourceURL": null,
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P018",
  "CategoryID": 1,
  "ProductName": "Nước tẩy quần áo màu Axo",
  "Manufacturer": "Công ty TNHH Green Cross Việt Nam",
  "SourceURL": "https://greencross.com.vn/nuoc-tay-quan-ao-mau",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid; C063 – Hydrogen peroxide; C109 – Sodium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P019",
  "CategoryID": 1,
  "ProductName": "Nước tẩy quần áo màu Hygiene Thái Lan",
  "Manufacturer": "Hygiene",
  "SourceURL": "https://hygiene.asia/thuoc-tay-moc-quan-ao-trang-mau-hygiene-thai-lan-250ml-sieu-sach-khong-gay-mui-tay",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C063 – Hydrogen peroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P020",
  "CategoryID": 1,
  "ProductName": "Nước xả vải khô Hygien SACOM Plus",
  "Manufacturer": "Hygiene",
  "SourceURL": "https://giaphanvn.com/nuoc-xa-vai-kho-hygiene",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C044 – Didecyldimethylammonium chloride",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P021",
  "CategoryID": 1,
  "ProductName": "Nước xịt vải Downy hương Huyền Bí (Downy Mystique Fabric Refresher Spray)",
  "Manufacturer": "Procter & Gamble (P&G)",
  "SourceURL": "https://www.lottemart.vn/vi-nsg/product/downy-mystique-fabric-refresher-spray-370ml-4987176112705-p45848?srsltid=AU7gw4V1BGIyO_70p31UxuJER6OKIN0VSaQFLbx2MUETAB2m6NDHgdyU",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C044 – Didecyldimethylammonium chloride",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P022",
  "CategoryID": 1,
  "ProductName": "Sản phẩm xịt tinh dầu thơm phòng và đuổi muỗi của thương hiệu Unitech.",
  "Manufacturer": "Công ty TNHH Unitech Việt Nam",
  "SourceURL": "https://www.lazada.vn/products/chai-xt-tinh-du-thm-phng--ui-mui-unitech-tinh-du-t-nhin-gip-ui-cn-trn-kh-khun-th-gin-ti-to-khng-kh-i3076412408.html?srsltid=AU7gw4UFQTD5wqE0lOgfmXETpRU33AJLY4ErQg8tudOO8C4lPj0NfSx3",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P023",
  "CategoryID": 1,
  "ProductName": "Thuốc tẩy trắng quần áo",
  "Manufacturer": "Chưa xác định",
  "SourceURL": null,
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C110 – Sodium hypochlorite",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P024",
  "CategoryID": 1,
  "ProductName": "Xịt phòng hương nước hoa Pure Air, thuộc thương hiệu Saigon Cosmetics Corporation (SCC).",
  "Manufacturer": "Saigon Cosmetics Corporation (SCC)",
  "SourceURL": "https://shop.saigoncosmetics.com/collections/xit-thom-phong",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol; C133 – Triethanolamine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P025",
  "CategoryID": 1,
  "ProductName": "Xịt thơm quần áo Yomosu hương nước hoa.",
  "Manufacturer": "Yomosu Cosmetics",
  "SourceURL": null,
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P026",
  "CategoryID": 2,
  "ProductName": "Dung dịch tạo bọt ngăn ngừa mụn Acnes Foaming Wash 150ml",
  "Manufacturer": "CÔNG TY TNHH ĐẦU TƯ THIÊN ÂN Đối Tác Phân Phối Ủy Quyền Của Công Ty TNHH Rohto - Mentholatum Việt Nam",
  "SourceURL": "https://shop.rohto.com.vn/acnes-foaming-wash.html?utm_term=acnes&gad_source=1&gad_campaignid=23556080761&gbraid=0AAAAA-_cMVnfGE9nvMG9qBfYq3QSp6QXM&gclid=CjwKCAjwtp7VBhBjEiwAJfpV-4sGlMZgsl86VP4DqNPBYlJl55QzvxL6m93B2EJUmffaIpz3eJhtmRoCCUQQAvD_BwE",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C082 – Niacinamide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P027",
  "CategoryID": 2,
  "ProductName": "Dầu gội Clear Men Cool Sport Bạc Hà.",
  "Manufacturer": "Unilever Việt Nam",
  "SourceURL": "https://www.clearhaircare.com/vn/p/dau-goi-clear-men-cool-sport-bac-ha.html/08934868132707",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 19,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C067 – Iodopropynyl butylcarbamate; C136 – Zinc pyrithione; C137 – Zinc sulfate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P028",
  "CategoryID": 2,
  "ProductName": "Dầu gội nước hoa X-Men For Boss Intense",
  "Manufacturer": "Công ty Cổ phần Hàng Gia dụng Quốc tế (ICP)",
  "SourceURL": "https://x-men.com.vn",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 11,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C079 – Methylchloroisothiazolinone; C080 – Methylisothiazolinone",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P029",
  "CategoryID": 2,
  "ProductName": "Dầu gội Romano Gen Z hương Sea Breeze.",
  "Manufacturer": "Công ty Cổ phần Hàng Gia dụng Quốc tế (ICP)",
  "SourceURL": "https://romano.vn",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 27,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C067 – Iodopropynyl butylcarbamate; C079 – Methylchloroisothiazolinone; C080 – Methylisothiazolinone",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P030",
  "CategoryID": 2,
  "ProductName": "Dầu gội đầu Rmano Attiude phiên bản 5 trong 1, loại 650g",
  "Manufacturer": "Công ty Cổ phần Hàng Gia dụng Quốc tế (ICP)",
  "SourceURL": "https://romano.vn",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 16,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C067 – Iodopropynyl butylcarbamate; C080 – Methylisothiazolinone; C136 – Zinc pyrithione; C137 – Zinc sulfate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P031",
  "CategoryID": 2,
  "ProductName": "Gel tẩy tế bào chết AWSOOYOR Sea Salt Extract Brightening and Decoloring Gel.",
  "Manufacturer": "AWSOOYOR Cosmetics",
  "SourceURL": "https://shopee.vn/AWSOOYOR-gel-t%E1%BA%A9y-t%E1%BA%BF-b%C3%A0o-ch%E1%BA%BFt-cho-m%E1%BA%B7t-v%C3%A0-c%C6%A1-th%E1%BB%83-th%E1%BA%A3o-m%E1%BB%99c-50g-l%C3%A0m-s%E1%BA%A1ch-d%E1%BB%8Bu-nh%E1%BA%B9-ki%E1%BB%83m-so%C3%A1t-d%E1%BA%A7u-da-v%E1%BA%ABn-%E1%BA%A9m-m%E1%BB%8Bn-i.1691704322.55003082819",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C128 – Tocopherol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P032",
  "CategoryID": 2,
  "ProductName": "Kem bôi da Hidem Pro Cream Hỗ trợ phục hồi da tổn thương do viêm da, chàm, vảy nến, kích ứng da, hăm tã và nấm da.",
  "Manufacturer": "Công ty Cổ phần Dược phẩm Gia Nguyễn",
  "SourceURL": "https://nhathuoclongchau.com.vn/thuoc/hidem-cream-15g-myung-in-pharm-14183.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 15,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C101 – Silver; C135 – Zinc oxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P033",
  "CategoryID": 2,
  "ProductName": "Kem bôi da Yoosun Rau Má (tuýp 25g) từ thương hiệu Đại Bắc, giúp làm mát, dịu da, ngăn ngừa rôm sảy, mẩn ngứa và hỗ trợ giảm mụn, mờ sẹo thâm.",
  "Manufacturer": "Công ty Cổ phần Dược phẩm BPPharma",
  "SourceURL": "Công ty Cổ phần Dược phẩm BPPharma",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 9,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C031 – Chlorhexidine digluconate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P034",
  "CategoryID": 2,
  "ProductName": "Kem đánh răng KINGBABY Fresh Breath Toothpaste, vị bạc hà, có khối lượng tịnh 100g",
  "Manufacturer": "CÔNG TY TNHH WONDERPLUS",
  "SourceURL": "https://www.the-kingbaby.com/h-pd-36.html",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 11,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C051 – DMDM hydantoin; C072 – Lysozyme; C076 – Menthol; C087 – Papain; ... (+2 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P035",
  "CategoryID": 2,
  "ProductName": "Lăn khử mùi nam X-Men For Boss Intense 50ml",
  "Manufacturer": "Công ty Cổ phần Hàng Gia dụng Quốc tế (ICP)",
  "SourceURL": "https://x-men.com.vn/xmen-for-boss/lan-khu-mui-intense/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol; C133 – Triethanolamine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P036",
  "CategoryID": 2,
  "ProductName": "Nước hoa nam BLUE DE Beautiful Girl",
  "Manufacturer": "Beautiful Girl Cosmetics",
  "SourceURL": "https://www.lazada.vn/products/nuoc-hoa-nam-nu-blue-de-beautyful-girl-cao-cap-i2379563260.html?srsltid=AU7gw4UFjGwy70JVCxZSrojFgyBqedbLyVMDOhQJl5APcOc2QfaGXTIs",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C035 – Citral; C041 – Coumarin; C053 – Ethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P037",
  "CategoryID": 2,
  "ProductName": "Nước hoa nam Dunté Bossy Men",
  "Manufacturer": "Dunté Perfume",
  "SourceURL": "https://nuochoaluahcm.com/san-pham/nuoc-hoa-unisex-lua-dante-50ml/",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P038",
  "CategoryID": 2,
  "ProductName": "Nước súc miệng kháng khuẩn Pearlie White Chlor-Rinse Plus.",
  "Manufacturer": "Pearlie White",
  "SourceURL": "https://nhathuoclongchau.com.vn/cham-soc-ca-nhan/nuoc-suc-mieng-sat-khuan-ngua-viem-nuou-chlor-rinse-plus-pearlie-white-250-ml.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 13,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C031 – Chlorhexidine digluconate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P039",
  "CategoryID": 2,
  "ProductName": "Nước súc miệng Listerine Cool Mint",
  "Manufacturer": "Công ty TNHH JNTL CONSUMER HEALTH (Việt Nam)",
  "SourceURL": "https://www.listerine.com.vn/mouthwash/antiseptic/listerine-cool-mint-mouthwash",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 8,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C016 – Benzoic acid; C053 – Ethanol; C056 – Eucalyptol; C076 – Menthol; ... (+2 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P040",
  "CategoryID": 2,
  "ProductName": "Nước súc miệng T-B Fresh của thương hiệu Traphaco, dung tích 500ml.",
  "Manufacturer": "T-B Fresh; Traphaco",
  "SourceURL": "https://nhathuoclongchau.com.vn/cham-soc-ca-nhan/nuoc-suc-mieng-sat-khuan-ngua-viem-nuou-chlor-rinse-plus-pearlie-white-250-ml.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 7,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C019 – Boric acid; C107 – Sodium fluoride",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P041",
  "CategoryID": 2,
  "ProductName": "Dầu gội dược liệu Nguyên Xuân Vàng dưỡng tóc, ngừa rụng tóc",
  "Manufacturer": "Công ty Cổ phần Dược phẩm Hoa Linh Hà Nam",
  "SourceURL": "https://duochoalinh.vn/dau-goi-duoc-lieu-nguyen-xuan-duong-toc/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 14,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C184 – Piroctone olamine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P042",
  "CategoryID": 2,
  "ProductName": "Dầu gội Clear 9 Thảo Dược Cổ Truyền",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": "https://u-shop.vn/d%E1%BA%A7u-g%E1%BB%99i-clear-9-th%E1%BA%A3o-d%C6%B0%E1%BB%A3c-c%E1%BB%95-truy%E1%BB%81n-630g",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 22,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C067 – Iodopropynyl butylcarbamate; C184 – Piroctone olamine",
  "ScopeNote": "Có chất chưa phân mức; đã có bằng chứng HPL-3 ở ít nhất một thành phần."
 },
 {
  "ProductID": "P043",
  "CategoryID": 2,
  "ProductName": "Dầu gội Sunsilk Óng Mượt Rạng Ngời",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam.",
  "SourceURL": "https://www.unilever.com.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 18,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C067 – Iodopropynyl butylcarbamate; C149 – BHT",
  "ScopeNote": "Có chất chưa phân mức; đã có bằng chứng HPL-3 ở ít nhất một thành phần."
 },
 {
  "ProductID": "P044",
  "CategoryID": 2,
  "ProductName": "Sữa tắm Lifebuoy Matcha & Khổ Qua",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": "https://www.unilever.com.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 24,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C184 – Piroctone olamine; C149 – BHT",
  "ScopeNote": "Có chất chưa phân mức; đã có bằng chứng HPL-3 ở ít nhất một thành phần."
 },
 {
  "ProductID": "P045",
  "CategoryID": 2,
  "ProductName": "Sữa tắm nước hoa Lux Botanicals sáng da hương hoa thiên điểu",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": "https://www.unilever.com.vn/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 8,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C112 – Sodium laureth sulfate; C040 – Cocamidopropyl betaine; C036 – Citric acid; C103 – Sodium benzoate; ... (+2 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P046",
  "CategoryID": 2,
  "ProductName": "Dầu gội phục hồi đa tầng nấm Truffle Trắng Weilaiya",
  "Manufacturer": "Weilaiya Việt Nam",
  "SourceURL": "https://hasaki.vn/thuong-hieu/weilaiya.html?srsltid=AfmBOoon90nrIYIqplhk2Il-it-B-uUnLlEfF-yMMpOKtDLxdOVWR5KG",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 10,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C080 – Methylisothiazolinone",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P047",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt Acnevir Santafa Face Wash Gel dạng tuýp 100g",
  "Manufacturer": "Santafa (Sản xuất bởi Công ty TNHH Dược Mỹ phẩm Santafa, Việt Nam).",
  "SourceURL": "https://santafa.vn/",
  "HPLProduct": "TỐI THIỂU HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 15,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C112 – Sodium laureth sulfate; C040 – Cocamidopropyl betaine; C068 – Lactic acid; C036 – Citric acid; ... (+5 chất cùng mức)",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P048",
  "CategoryID": 2,
  "ProductName": "Kem dưỡng ẩm và ngừa mụn ban đêm Canova Acnell Replenishing Cream",
  "Manufacturer": "Công ty TNHH Dược phẩm Biomedico",
  "SourceURL": "https://biomedico.vn/",
  "HPLProduct": "TỐI THIỂU HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 12,
  "NumUnrated": 2,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C024 – Caprylyl glycol; C089 – Phenoxyethanol; C103 – Sodium benzoate; C162 – Dimethicone",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P049",
  "CategoryID": 2,
  "ProductName": "Serium trị mụn Osaci Hana giúp giảm thâm, ngừa sẹo và dưỡng sáng da",
  "Manufacturer": "Osaci (Osaci Cosmetic) – đây là một thương hiệu mỹ phẩm nội địa được nghiên cứu và phát triển tại Việt Nam",
  "SourceURL": "https://osaci.vn/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 5,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C050 – Disodium EDTA; C054 – Ethylhexylglycerin; C089 – Phenoxyethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P050",
  "CategoryID": 2,
  "ProductName": "Serum dưỡng da mặt SADOER 5x Vitamin C Skin Active Face Serum",
  "Manufacturer": "SADOER – đây là một thương hiệu mỹ phẩm nội địa Trung Quốc (Chinese Skincare Brand)",
  "SourceURL": "https://osaci.vn/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P051",
  "CategoryID": 2,
  "ProductName": "Serum Ginseng New Skin, hỗ trợ các vấn đề về mụn, thâm, cải thiện tình trạng sẹo",
  "Manufacturer": "Công ty TNHH Thương mại và Dịch vụ Cao Minh Group",
  "SourceURL": "https://www.topcv.vn/cong-ty/cong-ty-tnhh-thuong-mai-va-dich-vu-cao-minh-group/182989.html",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 6,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C050 – Disodium EDTA; C068 – Lactic acid; C075 – Mandelic acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P052",
  "CategoryID": 2,
  "ProductName": "Sáp khử mùi Old Spice Pure Sport High Endurance",
  "Manufacturer": "Procter & Gamble (P&G)",
  "SourceURL": "https://echeck.numbala.com/cong-ty-tnhh-procter-gamble-viet-nam-dn81.htm",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C123 – Tetrasodium EDTA",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P053",
  "CategoryID": 2,
  "ProductName": "Sáp Vuốt Tóc X-Men For Boss Shine Pomade (loại 70g)",
  "Manufacturer": "Công ty Cổ phần Marico South East Asia sản xuất và phân phối",
  "SourceURL": "https://pnvt.vn/business-directory/cong-ty-co-phan-marico-south-east-asia/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 7,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol; C054 – Ethylhexylglycerin; C089 – Phenoxyethanol; C109 – Sodium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P054",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt +WIS+ Amino acid Soothing Cleanser",
  "Manufacturer": "WIS",
  "SourceURL": "https://metric.vn/ereport/thuong-hieu-wis?srsltid=AfmBOopxZ-8L7RL-HlbGC1napnTKupSBWWEdCCnuhEONrDUguii81dtw",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P055",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt -Facial Cleanser",
  "Manufacturer": null,
  "SourceURL": null,
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C112 – Sodium laureth sulfate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P056",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt AWSOOYOR Green Tea & Centella Cleanser",
  "Manufacturer": "AWSOOYOR Hàn Quốc",
  "SourceURL": "https://marketplace.tripmap.vn/articles/nguon-goc-va-do-tin-cay-cua-thuong-hieu-awsooyor",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid; C089 – Phenoxyethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P057",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt Awsooyor Hyaluronic Acid Cleanser.",
  "Manufacturer": "Awsooyor",
  "SourceURL": "https://marketplace.tripmap.vn/articles/nguon-goc-va-do-tin-cay-cua-thuong-hieu-awsooyor",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 9,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C001 – 1,2-Hexanediol; C054 – Ethylhexylglycerin; C089 – Phenoxyethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P058",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt dạng gel Miliman Version Pro Multi-Action, loại dành cho da dầu.",
  "Manufacturer": "MILIMAN",
  "SourceURL": "https://miliman.vn/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 12,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid; C040 – Cocamidopropyl betaine; C050 – Disodium EDTA; C051 – DMDM hydantoin; ... (+2 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P059",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt NIVEA MEN Acne Oil Clear Acne Defense Foam.",
  "Manufacturer": "Công ty TNHH Beiersdorf Việt Nam",
  "SourceURL": "https://www.fidovn.com/companies/20004877",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 18,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C012 – Arachidic acid; C053 – Ethanol; C068 – Lactic acid; C069 – Lauric acid; ... (+5 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P060",
  "CategoryID": 2,
  "ProductName": "Sữa rửa mặt The Cafuné Ginger Gel Cleanser",
  "Manufacturer": "Công ty Cổ phần Dược Mỹ Phẩm Cafuné",
  "SourceURL": "https://cafune.vn/",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C060 – Glycerol; C082 – Niacinamide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P061",
  "CategoryID": 2,
  "ProductName": "Sữa tắm cho nam Old Spice Fiji Body Wash với hương dừa và cây cọ",
  "Manufacturer": "Tập đoàn đa quốc gia Procter & Gamble (P&G) của Mỹ sản xuất",
  "SourceURL": "Không có đường link chính xác vì đây là sản phẩm sản xuất tại Mỹ",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 10,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C079 – Methylchloroisothiazolinone; C080 – Methylisothiazolinone",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P062",
  "CategoryID": 2,
  "ProductName": "Sữa tắm dưỡng ẩm, làm sáng da",
  "Manufacturer": null,
  "SourceURL": null,
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 6,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C040 – Cocamidopropyl betaine; C089 – Phenoxyethanol; C112 – Sodium laureth sulfate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P063",
  "CategoryID": 2,
  "ProductName": "Sữa tắm Verinice Red Wine",
  "Manufacturer": "VERINICE",
  "SourceURL": "https://cooponline.vn/sua-tam-verinice-ngua-lao-hoa-650ml--s250104870",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C036 – Citric acid; C040 – Cocamidopropyl betaine; C112 – Sodium laureth sulfate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P064",
  "CategoryID": 2,
  "ProductName": "Thuốc nhuộm tóc vĩnh viễn mang thương hiệu Purbrandy.",
  "Manufacturer": "Purbrandy",
  "SourceURL": null,
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 14,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C003 – 1-Naphthol; C011 – Ammonium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P065",
  "CategoryID": 2,
  "ProductName": "Tinh chất dưỡng da giảm mụn Cafuné Flux Derma Serum CARE+",
  "Manufacturer": "The Cafuné",
  "SourceURL": "https://cafune.vn/",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C060 – Glycerol; C082 – Niacinamide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P066",
  "CategoryID": 2,
  "ProductName": "Tinh chất dưỡng da Nuvessa",
  "Manufacturer": "Công ty Adamtino",
  "SourceURL": "https://www.topcv.vn/cong-ty/cong-ty-tnhh-thoi-trang-adamtino/252021.html",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C013 – Ascorbic acid; C082 – Niacinamide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P067",
  "CategoryID": 2,
  "ProductName": "Tinh chất dưỡng trắng Teashell Niacinamide Brightening Serum",
  "Manufacturer": "Teashell",
  "SourceURL": "https://teashell.com/products/tea-shell-niacinamide-brightening-serum",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C009 – Alpha-arbutin; C013 – Ascorbic acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P068",
  "CategoryID": 2,
  "ProductName": "Tinh chất Meso Peel White Serum giúp làm sáng da và tạo hiệu ứng căng bóng.",
  "Manufacturer": "Công ty TNHH Đầu tư - Thương mại Las Beauty The High",
  "SourceURL": "https://thehigh.vn/",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C082 – Niacinamide; C131 – Tranexamic acid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P069",
  "CategoryID": 2,
  "ProductName": "Tuýp kem nhuộm tóc vĩnh viễn thương hiệu Meiki của Nhật Bản, dung tích 100ml.",
  "Manufacturer": "Công ty TNHH Công nghệ Kaijoe Quảng Đông) có trụ sở tại tỉnh Guangdong, Trung Quốc.",
  "SourceURL": "Không có đường link chính xác vì đây là sản phẩm sản xuất tại Trung Quốc",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 5,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C011 – Ammonium hydroxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P070",
  "CategoryID": 2,
  "ProductName": "Tuýp thuốc nhuộm tóc chuyên nghiệp Gollagen (thường được biết đến là Roxy Collagen) dung tích 100ml, nổi bật với thành phần collagen giúp dưỡng tóc bóng mượt.",
  "Manufacturer": "Roxy Professional Việt Nam",
  "SourceURL": "http://www.roxyvietnam.com.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 6,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C011 – Ammonium hydroxide; C084 – p-Phenylenediamine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P071",
  "CategoryID": 2,
  "ProductName": "Xà phòng cục Lifebuoy Chăm Sóc Da với công thức Vitamin+.",
  "Manufacturer": "Công ty TNHH Quốc tế Unilever Việt Nam",
  "SourceURL": "https://www.unilever.com.vn/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 19,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C001 – 1,2-Hexanediol; C024 – Caprylyl glycol; C050 – Disodium EDTA; C069 – Lauric acid; ... (+7 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P072",
  "CategoryID": 2,
  "ProductName": "Xịt tinh dầu Antfamos của thương hiệu Medstand. Giúp làm thơm da, làm dịu da khi bị muỗi đốt, côn trùng cắn và góp phần làm mờ các vết thâm.",
  "Manufacturer": "Công ty Cổ phần Dược phẩm Medstand",
  "SourceURL": "https://trungtamthuoc.com/thuong-hieu/cong-ty-co-phan-duoc-pham-medstand",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C053 – Ethanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P073",
  "CategoryID": 3,
  "ProductName": "Bình xịt côn trùng ABS của BIGCO Group",
  "Manufacturer": "Công ty TNHH Bigco Group",
  "SourceURL": "https://bigcogroup.com.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C042 – Cypermethrin; C065 – Imiprothrin; C132 – Transfluthrin",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P074",
  "CategoryID": 3,
  "ProductName": "Bình xịt côn trùng Jumbo Vabe hương cam chanh",
  "Manufacturer": "Công ty TNHH Fumakilla Việt Nam",
  "SourceURL": "https://fumakilla.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C042 – Cypermethrin; C065 – Imiprothrin; C132 – Transfluthrin",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P075",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ cỏ Azinrio 45SC đặc trị cỏ ngô, mía, quế. Thể tích 80ml",
  "Manufacturer": "Công ty TNHH XNK hóa chất CNC Việt Nam",
  "SourceURL": "https://masothue.com/0109776898-cong-ty-tnhh-xuat-nhap-khau-hoa-chat-cong-nghe-cao-viet-nam",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C015 – Atrazine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P076",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ cỏ BASTNATE 200SL Sản xuất bởi Trung Quốc",
  "Manufacturer": "Rainbow Agro (Shandong Weifang Rainbow Chemical Co., Ltd.)",
  "SourceURL": "https://www.rainbowagro.com/gl",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C058 – Glufosinate-ammonium",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P077",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ cỏ Newfosinate 150SL, nhãn hiệu Trâu Rừng mới.",
  "Manufacturer": "Công ty CP BMC Việt Nam",
  "SourceURL": "https://bmcgroup.com.vn/vi/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C058 – Glufosinate-ammonium",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P078",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu Acnal 400WP nhãn hiệu rồng lưở Gold chuyên trị rầy xanh, rầy nâu",
  "Manufacturer": "Công ty TNHH Thương mại Thái Nông",
  "SourceURL": "https://www.yellowpages.vn/lgs/695166/cong-ty-tnhh-thuong-mai-thai-nong.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C007 – Acetamiprid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P079",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu BILLADEN 50WG nhãn hiệu Rồng Vàng",
  "Manufacturer": "Công ty Cổ phần Nông nghiệp Thanh Xuân. (Có trụ sở đăng ký tại Hóc Môn, TP. Hồ Chí Minh)",
  "SourceURL": "https://dauthau.net/vi/dn/CONG-TY-CO-PHAN-NONG-NGHIEP-THANH-XUAN/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C052 – Emamectin benzoate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P080",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu Biperin 100EC chuyên trị các loại sâu hại lúa và cây trồng khác",
  "Manufacturer": "Công ty TNHH Thương mại Bình Phương.",
  "SourceURL": "https://binhphuongco.com/lien-he.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C042 – Cypermethrin",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P081",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu Chipro 550 nhãn hiệu Rồng Vàng 68, 400ml",
  "Manufacturer": "Công ty cổ phần Hùng Bình",
  "SourceURL": "https://fiingate.vn/BusinessDirectory/CompanyDetail/1987194?lang=vi-VN",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C092 – Profenofos; C125 – Thiamethoxam",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P082",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu forgon 40EC nhãn hiệu Bowing 777, đặc trị sâu cuốn lá, rầy xanh, rệp sáp.",
  "Manufacturer": "Công ty cổ phần thương mại Hải Ánh",
  "SourceURL": "https://hacoagri.com.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C046 – Dimethoate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P083",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu Rồng vàng 9999, đặc trị sâu đục thân.",
  "Manufacturer": "Công ty Cổ phần Nông nghiệp Thiên Minh",
  "SourceURL": "https://thienminhagri.com/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C025 – Carbosulfan",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P084",
  "CategoryID": 3,
  "ProductName": "Thuốc trừ sâu, nhãn hiệu: ACTADOR 100WP",
  "Manufacturer": "Công ty CP công nghệ cao thuốc BVtV USA",
  "SourceURL": "https://usachemical.vn/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C064 – Imidacloprid",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P085",
  "CategoryID": 3,
  "ProductName": "Chai thuốc xịt diệt côn trùng An Phát, loại 350ml, có công dụng diệt ruồi, muỗi, kiến và gián.",
  "Manufacturer": "Công ty TNHH Một thành viên Thương mại Sản xuất An Phát",
  "SourceURL": "https://www.yellowpages.vn/lgs/1187875116/diet-con-trung-an-phat-cong-ty-tnhh-mot-thanh-vien-thuong-mai-san-xuat-an-phat.html",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C010 – Alpha-cypermethrin; C043 – Deltamethrin",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P086",
  "CategoryID": 4,
  "ProductName": "Kem làm bánh Rich's Whip Topping Base 907 g",
  "Manufacturer": "Rich's / Rich Products Vietnam",
  "SourceURL": "https://richs.com.vn/whip-topping/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 10,
  "NumUnrated": 7,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C105 – Sodium chloride; C106 – Sodium citrate; C198 – Xanthan gum",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P087",
  "CategoryID": 4,
  "ProductName": "Kem trang trí bánh Rich's Nhãn Vàng Gold Label 907 g",
  "Manufacturer": "Rich's / Rich Products Vietnam",
  "SourceURL": "https://shop.richs.com.vn/nguyen-lieu-lam-banh/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 7,
  "NumUnrated": 6,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C105 – Sodium chloride",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P088",
  "CategoryID": 4,
  "ProductName": "Kem sữa trang trí và làm bánh Rich's Niagara Farms 1 kg",
  "Manufacturer": "Rich's / Rich Products Vietnam",
  "SourceURL": "https://richs.com.vn/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 12,
  "NumUnrated": 8,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C105 – Sodium chloride; C198 – Xanthan gum; C170 – Guar gum (E412); C153 – Beta-carotene (E160a(i))",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P089",
  "CategoryID": 4,
  "ProductName": "Bột bánh bông lan Richmix 1 kg",
  "Manufacturer": "Richmix / Rich Products Vietnam",
  "SourceURL": "https://shop.richs.com.vn/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 11,
  "NumUnrated": 8,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C105 – Sodium chloride; C192 – Sodium bicarbonate; C198 – Xanthan gum",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P090",
  "CategoryID": 4,
  "ProductName": "Bột bánh Muffin Richmix vị Vanilla 1 kg",
  "Manufacturer": "Richmix / Rich Products Vietnam",
  "SourceURL": "https://shop.richs.com.vn/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 9,
  "NumUnrated": 7,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C105 – Sodium chloride; C192 – Sodium bicarbonate",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P091",
  "CategoryID": 4,
  "ProductName": "Bột rau câu dẻo Jelly Rovin",
  "Manufacturer": "Rovin / Công ty TNHH Rovin",
  "SourceURL": "https://rovinco.vn/cua-hang/bot-rau-cau-deo-rovin/",
  "HPLProduct": "TỐI THIỂU HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 3,
  "NumUnrated": 2,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C188 – Potassium chloride",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P092",
  "CategoryID": 4,
  "ProductName": "Bột rau câu giòn Agar Rovin 25 g",
  "Manufacturer": "Rovin / Công ty TNHH Rovin",
  "SourceURL": "https://rovinco.vn/",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C147 – Agar-agar",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P093",
  "CategoryID": 4,
  "ProductName": "Hương Vanilla Rayner's",
  "Manufacturer": "Rayner's / HOLAFOODS",
  "SourceURL": "https://holafoods.com.vn/blogs/ho-so-cong-bo-san-pham/ho-so-cong-bo-san-pham",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P094",
  "CategoryID": 4,
  "ProductName": "Màu thực phẩm đỏ Rayner's",
  "Manufacturer": "Rayner's / HOLAFOODS",
  "SourceURL": "https://holafoods.com.vn/blogs/ho-so-cong-bo-san-pham/ho-so-cong-bo-san-pham",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P095",
  "CategoryID": 4,
  "ProductName": "Màu thực phẩm đỏ Wilton",
  "Manufacturer": "Wilton / HOLAFOODS",
  "SourceURL": "https://holafoods.com.vn/blogs/ho-so-cong-bo-san-pham/ho-so-cong-bo-san-pham",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P096",
  "CategoryID": 5,
  "ProductName": "Bình xịt bôi trơn đa năng Selleys RP7",
  "Manufacturer": "Tập đoàn sơn và chất phủ toàn cầu Nippon Paint",
  "SourceURL": "https://nipponpaint.com.vn/vi",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C095 – Propanol",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P097",
  "CategoryID": 5,
  "ProductName": "Sơn chống thấm: Jotun WaterGuard",
  "Manufacturer": "Công ty TNHH Sơn Jotun Việt Nam.",
  "SourceURL": "https://www.jotun.com/vn-vi/decorative",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P098",
  "CategoryID": 5,
  "ProductName": "Sơn ngoại thất: Jotun Tough Shield",
  "Manufacturer": "Công ty TNHH Sơn Jotun Việt Nam.",
  "SourceURL": "https://www.jotun.com/vn-vi/decorative",
  "HPLProduct": "CHƯA XÁC ĐỊNH",
  "HPLScoreMax": 0,
  "NumChemicals": 0,
  "NumUnrated": 0,
  "AssessmentStatus": "THIẾU THÀNH PHẦN",
  "DominantChemicals": null,
  "ScopeNote": "Chưa xác định thành phần; không được tự gán HPL."
 },
 {
  "ProductID": "P099",
  "CategoryID": 5,
  "ProductName": "Sơn nội thất: Jotun Jotaplast",
  "Manufacturer": "Công ty TNHH Sơn Jotun Việt Nam.",
  "SourceURL": "https://www.jotun.com/vn-vi/decorative",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C127 – Titanium dioxide",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P100",
  "CategoryID": 5,
  "ProductName": "Bình xịt đa năng WD-40 Multi-Use Product",
  "Manufacturer": "WD-40",
  "SourceURL": "https://wd40.asia/vn/product/wd-40-multi-use-product/",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C183 – Petroleum distillates, hydrotrea",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P101",
  "CategoryID": 5,
  "ProductName": "Chai xịt vệ sinh tiếp điểm WD-40 Specialist Contact Cleaner",
  "Manufacturer": "WD-40",
  "SourceURL": "https://wd40.asia/vn/product/",
  "HPLProduct": "TỐI THIỂU HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 3,
  "NumUnrated": 1,
  "AssessmentStatus": "CHƯA ĐỦ BẰNG CHỨNG",
  "DominantChemicals": "C199 – n-Pentane",
  "ScopeNote": "Có chất chưa phân mức: HPL chỉ là mức tối thiểu trên các chất đã đánh giá."
 },
 {
  "ProductID": "P102",
  "CategoryID": 5,
  "ProductName": "Chai xịt vệ sinh máy lạnh WD-40 3-IN-ONE Air Conditioner Cleaner",
  "Manufacturer": "WD-40 / 3-IN-ONE",
  "SourceURL": "https://wd40.asia/vn/products/diy/",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C150 – Benzalkonium chloride (C12–C16 a",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P103",
  "CategoryID": 5,
  "ProductName": "Keo dán xây dựng Selleys Liquid Nails Original",
  "Manufacturer": "Selleys / Nippon Paint Vietnam",
  "SourceURL": "https://go.lupinsys.com/duluxgroup/harms/public/materials/b55580851bf80bd917f9509927249a8c-published/attachments_api/8beb34e9a44d27b5dc41725addb6f454/search_api/SELLEYS_LIQUID_NAILS-VNM_SDS.pdf",
  "HPLProduct": "HPL-3",
  "HPLScoreMax": 3,
  "NumChemicals": 2,
  "NumUnrated": 0,
  "AssessmentStatus": "CẦN XÁC MINH THÀNH PHẦN",
  "DominantChemicals": "C144 – 2,2'-Methylenebis(6-tert-butyl-4",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp. CẦN XÁC MINH NGUỒN THÀNH PHẦN: C182: SDS Selleys Liquid Nails Original phiên bản NZ 13/03/2024; cần xác nhận phiên bản bán tại Việt Nam.; C144: SDS Selleys Liquid Nails Original phiên bản NZ 13/03/2024; cần xác nhận phiên bản bán tại Việt Nam."
 },
 {
  "ProductID": "P104",
  "CategoryID": 5,
  "ProductName": "Keo trám khe Selleys No More Gaps",
  "Manufacturer": "Selleys",
  "SourceURL": "https://tinhiep.com/san-pham/keo-tram-khe-selleys-no-more-gaps/",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 1,
  "NumUnrated": 0,
  "AssessmentStatus": "CẦN XÁC MINH THÀNH PHẦN",
  "DominantChemicals": "C022 – Calcium carbonate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp. CẦN XÁC MINH NGUỒN THÀNH PHẦN: C022: SDS No More Gaps Multipurpose đời cũ nêu Calcium carbonate 30–60%; cần xác nhận biến thể và nhãn bán tại Việt Nam."
 },
 {
  "ProductID": "P105",
  "CategoryID": 5,
  "ProductName": "Dầu phanh Castrol Brake Fluid DOT 4",
  "Manufacturer": "Castrol",
  "SourceURL": "https://www.castrol.com/vi_vn/vietnam/home/products/automotive/automotive-engines/brake-fluids.html",
  "HPLProduct": "HPL-2",
  "HPLScoreMax": 2,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C045 – Diethylene glycol; C197 – Tris[2-(2-(2-methoxyethoxy)ethox; C145 – 2-[2-(2-Butoxyethoxy)ethoxy]etha; C161 – Diisopropanolamine",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P106",
  "CategoryID": 4,
  "ProductName": "Heo 2 Lát VISSAN 150 g",
  "Manufacturer": "VISSAN / Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản",
  "SourceURL": "https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/heo-2-lat.html",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 4,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C138 – Potassium nitrite; C140 – Sodium tripolyphosphate; C141 – Sodium polyphosphate; C142 – Monosodium glutamate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P107",
  "CategoryID": 4,
  "ProductName": "Pâté Thịt Heo VISSAN 170 g",
  "Manufacturer": "VISSAN / Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản",
  "SourceURL": "https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/pate-thit-heo.html",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 6,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C138 – Potassium nitrite; C139 – Sodium erythorbate; C140 – Sodium tripolyphosphate; C141 – Sodium polyphosphate; ... (+2 chất cùng mức)",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 },
 {
  "ProductID": "P108",
  "CategoryID": 4,
  "ProductName": "Pâté Thịt Bò Đặc Biệt VISSAN 170 g",
  "Manufacturer": "VISSAN / Công ty Cổ phần Việt Nam Kỹ nghệ Súc sản",
  "SourceURL": "https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/pate-thit-bo-dac-biet.html",
  "HPLProduct": "HPL-1",
  "HPLScoreMax": 1,
  "NumChemicals": 3,
  "NumUnrated": 0,
  "AssessmentStatus": "ĐỦ HPL CHO CÁC CHẤT ĐÃ GẮN",
  "DominantChemicals": "C138 – Potassium nitrite; C139 – Sodium erythorbate; C142 – Monosodium glutamate",
  "ScopeNote": "Theo mức HPL cao nhất trong các thành phần đã xác minh; không phải đánh giá rủi ro hỗn hợp."
 }
];
function getProductById(id) { return PRODUCTS.find(p => p.ProductID === id) || null; }


/* ═══════════════════════════════════════════════════════════════════════
   6. CHEMICAL — 199 hồ sơ hóa chất (C001–C143 gốc; C144–C199 mới)
   ═══════════════════════════════════════════════════════════════════════ */
const CHEMICALS = [
 {
  "ChemicalID": "C001",
  "IsNew": false,
  "ChemicalName": "1,2-Hexanediol",
  "Synonyms": [
   "Hexane-1,2-diol"
  ],
  "CASNumber": "6920-22-5",
  "Formula": "C6H14O2",
  "ChemicalGroup": "Diol / chất giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation (supplier-dependent)",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "1450 (cá, 96 h; ECOSAR ước tính)",
   "EC50": "330 (tảo, 96 h; ECOSAR ước tính)",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "0,69 (ước tính; giá trị thực nghiệm đầu vào ~0,58)",
   "BCF": "3,2 (ước tính)",
   "Biodegradation": "82,1–83,7%/28 ngày, OECD 301B → dễ phân hủy sinh học",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Thấp hơn về độc tính cấp trong endpoint hiện có; vẫn cần xem xét dữ liệu mạn và đặc tính môi trường",
   "DataQuality": "QSAR + thử nghiệm biodegradation",
   "EcoSource": "PubChem/EPA-OECD screening; tra cứu CAS 6920-22-5 | https://pubchem.ncbi.nlm.nih.gov/#query=6920-22-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=6920-22-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C002",
  "IsNew": false,
  "ChemicalName": "1-Hydroxyethyl-4,5-diamino pyrazole sulfate",
  "Synonyms": [
   "1-Hydroxyethyl-4,5-diaminopyrazole sulfate"
  ],
  "CASNumber": "155601-30-2",
  "Formula": "C5H12N4O5S",
  "ChemicalGroup": "Pyrazole hair dye intermediate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin sensitization; eye irritation (supplier-dependent)",
  "HazardCategories": "1; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H317",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây phản ứng dị ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=155601-30-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H317, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=155601-30-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C003",
  "IsNew": false,
  "ChemicalName": "1-Naphthol",
  "Synonyms": [
   "alpha-Naphthol"
  ],
  "CASNumber": "90-15-3",
  "Formula": "C10H8O",
  "ChemicalGroup": "Naphthol / hair dye coupler",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin/eye irritation; Aquatic hazard",
  "HazardCategories": "4; 2; 2; 1",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H315",
   "H319",
   "H410"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "4,24 (Pimephales promelas, 96 h; thực nghiệm)",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "2,85 (thực nghiệm)",
   "BCF": "31 (ước tính)",
   "Biodegradation": "96% BOD/2 tuần (MITI); đạt tiêu chí dễ phân hủy",
   "AquaticClass": "Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm + ước tính BCF",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/1-Naphthol"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=90-15-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C004",
  "IsNew": false,
  "ChemicalName": "2-Amino-6-chloro-4-nitrophenol",
  "Synonyms": [
   "2-Amino-6-chloro-4-nitrophenol"
  ],
  "CASNumber": "6358-09-4",
  "Formula": "C6H5ClN2O3",
  "ChemicalGroup": "Nitrophenol dye intermediate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Cần đối chiếu ECHA/SDS; có nguy cơ kích ứng/nhạy cảm",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=6358-09-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=6358-09-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C005",
  "IsNew": false,
  "ChemicalName": "2-Methylresorcinol",
  "Synonyms": [
   "2-Methylbenzene-1,3-diol"
  ],
  "CASNumber": "608-25-3",
  "Formula": "C7H8O2",
  "ChemicalGroup": "Phenol / hair dye coupler",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin/eye irritation; skin sensitization possible",
  "HazardCategories": "2; 2; 1",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H315",
   "H317",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có thể gây phản ứng dị ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=608-25-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H317, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=608-25-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C006",
  "IsNew": false,
  "ChemicalName": "Acesulfame potassium",
  "Synonyms": [
   "Acesulfame K"
  ],
  "CASNumber": "55589-62-3",
  "Formula": "C4H4KNO4S",
  "ChemicalGroup": "Artificial sweetener",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=55589-62-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=55589-62-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C007",
  "IsNew": false,
  "ChemicalName": "Acetamiprid",
  "Synonyms": [
   "Acetamiprid"
  ],
  "CASNumber": "135410-20-7",
  "Formula": "C10H11ClN4",
  "ChemicalGroup": "Neonicotinoid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity",
  "HazardCategories": "4",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302"
  ],
  "HealthEffects": "Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "100 (Cyprinodon variegatus, 96 h; thực nghiệm)",
   "EC50": "49,8 (Daphnia, 48 h; thực nghiệm)",
   "NOEC": "0,0025 (mysid, 28 ngày, NOEC); EC₁₀ 0,000235 (Chironomus; mạn)",
   "logKow": "0,8",
   "BCF": "2 (ước tính)",
   "Biodegradation": "27%/28 ngày (OECD 301B) → không dễ phân hủy",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Cao – độc tính mạn ở nồng độ rất thấp kết hợp khả năng phân hủy hạn chế/chậm",
   "DataQuality": "Thực nghiệm + ước tính BCF",
   "EcoSource": "EFSA/EPA data via authoritative regulatory records; CAS 135410-20-7 | https://pubchem.ncbi.nlm.nih.gov/#query=135410-20-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Độc tính mạn ≤0,1 mg/L và không dễ phân hủy (giá trị thấp nhất 0.000235 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=135410-20-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C008",
  "IsNew": false,
  "ChemicalName": "Allantoin",
  "Synonyms": [
   "Glyoxyldiureide"
  ],
  "CASNumber": "97-59-6",
  "Formula": "C4H6N4O3",
  "ChemicalGroup": "Urea derivative / skin protectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "−3,14 (ước tính)",
   "BCF": "3 (ước tính)",
   "Biodegradation": "88%/30 ngày (closed-bottle; tăng từ 6% ngày 5 và 10% ngày 15)",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng để xếp mức sinh thái độc lập; sử dụng GHS và dữ liệu fate làm căn cứ bổ trợ",
   "DataQuality": "Ước tính + thử nghiệm biodegradation",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=97-59-6"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=97-59-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C009",
  "IsNew": false,
  "ChemicalName": "Alpha-arbutin",
  "Synonyms": [
   "4-Hydroxyphenyl alpha-D-glucopyranoside"
  ],
  "CASNumber": "84380-01-8",
  "Formula": "C12H16O7",
  "ChemicalGroup": "Glycoside / skin-lightening agent",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=84380-01-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=84380-01-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C010",
  "IsNew": false,
  "ChemicalName": "Alpha-cypermethrin",
  "Synonyms": [
   "α-Cypermethrin"
  ],
  "CASNumber": "67375-30-8",
  "Formula": "C22H19Cl2NO3",
  "ChemicalGroup": "Pyrethroid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; STOT; Aquatic acute/chronic",
  "HazardCategories": "3/4; 2; 1/1",
  "Pictograms": [
   "GHS06",
   " GHS08",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301/H332",
   "H373",
   "H400",
   "H410"
  ],
  "HealthEffects": "Độc/nguy hiểm khi nuốt phải. Có hại khi hít phải. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "0,0028 (Oncorhynchus mykiss, 96 h; thực nghiệm)",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "6,94",
   "BCF": "990 (ước tính)",
   "Biodegradation": "Chuyển hóa sinh học hiếu khí; t½ khoảng 7–14 ngày trong hệ vi sinh thích nghi; không coi là ready-biodegradable từ dữ liệu này",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm + ước tính",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=67375-30-8"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1; Độc tính cấp thủy sinh đáng tin cậy ≤1 mg/L (giá trị thấp nhất 0.0028 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=67375-30-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C011",
  "IsNew": false,
  "ChemicalName": "Ammonium hydroxide",
  "Synonyms": [
   "Aqueous ammonia",
   "Ammonia solution"
  ],
  "CASNumber": "1336-21-6",
  "Formula": "NH4OH (aq)",
  "ChemicalGroup": "Base vô cơ / dung dịch amoniac",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin corrosion; Aquatic acute",
  "HazardCategories": "1B; 1",
  "Pictograms": [
   "GHS05",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H314",
   "H400"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "0,024–0,093 (Lepomis macrochirus, 48 h; phụ thuộc pH/speciation)",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Aquatic Acute 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm; chất vô cơ/ion hóa",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1336-21-6"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1; Độc tính cấp thủy sinh đáng tin cậy ≤1 mg/L (giá trị thấp nhất 0.024 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1336-21-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C012",
  "IsNew": false,
  "ChemicalName": "Arachidic acid",
  "Synonyms": [
   "Eicosanoic acid"
  ],
  "CASNumber": "506-30-9",
  "Formula": "C20H40O2",
  "ChemicalGroup": "Fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "4,67 (cá, 96 h; QSAR)",
   "EC50": "1,22 (Daphnia, 48 h; QSAR); 3,806 (tảo, 96 h; QSAR)",
   "NOEC": "5 (tảo, 72 h; read-across, không phải chất đích)",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "ECHA kết luận không đáng lo về tích lũy sau xem xét BCF/biodegradation; chưa trích được giá trị đơn nhất phù hợp",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Trung bình – endpoint độc tính cấp nằm trong khoảng >1–100 mg/L",
   "DataQuality": "QSAR/read-across – độ tin cậy trung bình",
   "EcoSource": "ECHA REACH dossier, CAS 506-30-9 | https://pubchem.ncbi.nlm.nih.gov/#query=506-30-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Độc tính cấp thủy sinh >1–100 mg/L (giá trị thấp nhất 1.22 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=506-30-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C013",
  "IsNew": false,
  "ChemicalName": "Ascorbic acid",
  "Synonyms": [
   "Vitamin C"
  ],
  "CASNumber": "50-81-7",
  "Formula": "C6H8O6",
  "ChemicalGroup": "Vitamin / antioxidant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "1020 (Oncorhynchus mykiss, 96 h; thực nghiệm)",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "−1,85 (thực nghiệm)",
   "BCF": "~3,16 (ước tính từ logBCF 0,5)",
   "Biodegradation": "Dự báo phân hủy cuối cùng t½ ~8,67 ngày; xác suất MITI cao",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Thấp hơn về độc tính cấp trong endpoint hiện có; vẫn cần xem xét dữ liệu mạn và đặc tính môi trường",
   "DataQuality": "Thực nghiệm + mô hình",
   "EcoSource": "OECD/Canadian DSL & PubChem; CAS 50-81-7 | https://pubchem.ncbi.nlm.nih.gov/#query=50-81-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=50-81-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C014",
  "IsNew": false,
  "ChemicalName": "Asiaticoside",
  "Synonyms": [
   "Asiaticoside"
  ],
  "CASNumber": "16830-15-2",
  "Formula": "C48H78O19",
  "ChemicalGroup": "Triterpenoid glycoside",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không có phân loại GHS hài hòa phổ biến; cần SDS nguyên liệu",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=16830-15-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=16830-15-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C015",
  "IsNew": false,
  "ChemicalName": "Atrazine",
  "Synonyms": [
   "Atrazine"
  ],
  "CASNumber": "1912-24-9",
  "Formula": "C8H14ClN5",
  "ChemicalGroup": "Triazine herbicide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Skin sensitization; STOT RE; Aquatic chronic",
  "HazardCategories": "1; 2; 1",
  "Pictograms": [
   "GHS07",
   " GHS08",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H317",
   "H373",
   "H410"
  ],
  "HealthEffects": "Có thể gây phản ứng dị ứng da. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "2,61",
   "BCF": "<0,27–96 (cá; khoảng báo cáo)",
   "Biodegradation": "Phân hủy chậm/biến thiên; t½ đất thường ~28–35 ngày, có thể kéo dài hơn trong nước ngầm",
   "AquaticClass": "Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm/field; endpoint acute chưa chuẩn hóa",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1912-24-9"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1912-24-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C016",
  "IsNew": false,
  "ChemicalName": "Benzoic acid",
  "Synonyms": [
   "Benzenecarboxylic acid"
  ],
  "CASNumber": "65-85-0",
  "Formula": "C7H6O2",
  "ChemicalGroup": "Aromatic carboxylic acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin irritation; Serious eye damage",
  "HazardCategories": "2; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H315",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây kích ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=65-85-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=65-85-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C017",
  "IsNew": false,
  "ChemicalName": "Benzyl alcohol",
  "Synonyms": [
   "Phenylmethanol"
  ],
  "CASNumber": "100-51-6",
  "Formula": "C7H8O",
  "ChemicalGroup": "Aromatic alcohol",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Eye irritation",
  "HazardCategories": "4; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H332",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có hại khi nuốt phải. Có hại khi hít phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "205 (cá; QSAR/ước tính)",
   "EC50": "55 (Daphnia, 24 h; thực nghiệm)",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "1,1 (thực nghiệm)",
   "BCF": "~0,31 (ước tính)",
   "Biodegradation": "96% → dễ phân hủy sinh học",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Trung bình – endpoint độc tính cấp nằm trong khoảng >1–100 mg/L",
   "DataQuality": "Thực nghiệm + QSAR",
   "EcoSource": "OECD/Canadian DSL; CAS 100-51-6 | https://pubchem.ncbi.nlm.nih.gov/#query=100-51-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H319, H332; Độc tính cấp thủy sinh >1–100 mg/L (giá trị thấp nhất 55 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=100-51-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C018",
  "IsNew": false,
  "ChemicalName": "Biotin",
  "Synonyms": [
   "Vitamin B7"
  ],
  "CASNumber": "58-85-5",
  "Formula": "C10H16N2O3S",
  "ChemicalGroup": "Vitamin",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "201,1 (cá; QSAR)",
   "EC50": "8284 (ECOSAR, ước tính)",
   "NOEC": "1435,96 (chronic value dự báo; không phải NOEC thực nghiệm)",
   "logKow": "0,39 (ước tính)",
   "BCF": "~3,16 (ước tính)",
   "Biodegradation": "Dự báo t½ phân hủy cuối cùng ~15 ngày; bằng chứng ready-biodegradation chưa chắc chắn",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Thấp hơn về độc tính cấp trong endpoint hiện có; vẫn cần xem xét dữ liệu mạn và đặc tính môi trường",
   "DataQuality": "QSAR/ước tính",
   "EcoSource": "OECD/Canadian DSL; CAS 58-85-5 | https://pubchem.ncbi.nlm.nih.gov/#query=58-85-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=58-85-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C019",
  "IsNew": false,
  "ChemicalName": "Boric acid",
  "Synonyms": [
   "Orthoboric acid"
  ],
  "CASNumber": "10043-35-3",
  "Formula": "H3BO3",
  "ChemicalGroup": "Acid vô cơ yếu",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Reproductive toxicity",
  "HazardCategories": "1B",
  "Pictograms": [
   "GHS08"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H360FD"
  ],
  "HealthEffects": "Có thể gây tổn hại khả năng sinh sản hoặc thai nhi.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=10043-35-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H360",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=10043-35-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C020",
  "IsNew": false,
  "ChemicalName": "Butane",
  "Synonyms": [
   "n-Butane"
  ],
  "CASNumber": "106-97-8",
  "Formula": "C4H10",
  "ChemicalGroup": "Hydrocarbon / propellant",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Flammable gas; Gas under pressure",
  "HazardCategories": "1; liquefied gas",
  "Pictograms": [
   "GHS02",
   " GHS04"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H220",
   "H280"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=106-97-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=106-97-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C021",
  "IsNew": false,
  "ChemicalName": "Butylene glycol",
  "Synonyms": [
   "1,3-Butanediol"
  ],
  "CASNumber": "107-88-0",
  "Formula": "C4H10O2",
  "ChemicalGroup": "Diol / dung môi giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại theo GHS",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=107-88-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=107-88-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C022",
  "IsNew": false,
  "ChemicalName": "Calcium carbonate",
  "Synonyms": [
   "Calcite",
   "Limestone"
  ],
  "CASNumber": "471-34-1",
  "Formula": "CaCO3",
  "ChemicalGroup": "Muối carbonate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=471-34-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=471-34-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C023",
  "IsNew": false,
  "ChemicalName": "Calcium pantothenate",
  "Synonyms": [
   "Calcium D-pantothenate"
  ],
  "CASNumber": "137-08-6",
  "Formula": "C18H32CaN2O10",
  "ChemicalGroup": "Vitamin salt",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=137-08-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=137-08-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C024",
  "IsNew": false,
  "ChemicalName": "Caprylyl glycol",
  "Synonyms": [
   "1,2-Octanediol"
  ],
  "CASNumber": "1117-86-8",
  "Formula": "C8H18O2",
  "ChemicalGroup": "Diol / humectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation possible",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1117-86-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1117-86-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C025",
  "IsNew": false,
  "ChemicalName": "Carbosulfan",
  "Synonyms": [
   "Carbosulfan"
  ],
  "CASNumber": "55285-14-8",
  "Formula": "C20H32N2O3S",
  "ChemicalGroup": "Carbamate insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic acute/chronic",
  "HazardCategories": "3; 1/1",
  "Pictograms": [
   "GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H400",
   "H410"
  ],
  "HealthEffects": "Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=55285-14-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=55285-14-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C026",
  "IsNew": false,
  "ChemicalName": "Carnitine",
  "Synonyms": [
   "L-Carnitine"
  ],
  "CASNumber": "541-15-1",
  "Formula": "C7H15NO3",
  "ChemicalGroup": "Quaternary ammonium metabolite",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=541-15-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=541-15-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C027",
  "IsNew": false,
  "ChemicalName": "Cetearyl alcohol",
  "Synonyms": [
   "Cetostearyl alcohol"
  ],
  "CASNumber": "67762-27-0",
  "Formula": "Mixture C16H34O/C18H38O",
  "ChemicalGroup": "Fatty alcohol mixture",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại đáng kể; đây là hỗn hợp alcohol béo xác định thương mại",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=67762-27-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=67762-27-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C028",
  "IsNew": false,
  "ChemicalName": "Cetrimonium chloride",
  "Synonyms": [
   "Hexadecyltrimethylammonium chloride"
  ],
  "CASNumber": "112-02-7",
  "Formula": "C19H42ClN",
  "ChemicalGroup": "Quaternary ammonium compound",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Acute toxicity; Skin/eye damage; Aquatic hazard",
  "HazardCategories": "4; 2/1; 1",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H315",
   "H318",
   "H400"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây kích ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "0,067 (Brachionus calyciflorus, 48 h; thực nghiệm, không phải Daphnia)",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "3,23 (thực nghiệm)",
   "BCF": "~70,8 (ước tính)",
   "Biodegradation": "70,3% → có khả năng dễ phân hủy theo thử nghiệm được báo cáo",
   "AquaticClass": "Aquatic Acute 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm + ước tính BCF",
   "EcoSource": "OECD/Canadian screening; CAS 112-02-7 | https://pubchem.ncbi.nlm.nih.gov/#query=112-02-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1; Độc tính cấp thủy sinh đáng tin cậy ≤1 mg/L (giá trị thấp nhất 0.067 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=112-02-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C029",
  "IsNew": false,
  "ChemicalName": "Cetyl alcohol",
  "Synonyms": [
   "1-Hexadecanol"
  ],
  "CASNumber": "36653-82-4",
  "Formula": "C16H34O",
  "ChemicalGroup": "Fatty alcohol",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=36653-82-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=36653-82-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C030",
  "IsNew": false,
  "ChemicalName": "Cetylpyridinium chloride",
  "Synonyms": [
   "CPC"
  ],
  "CASNumber": "123-03-5",
  "Formula": "C21H38ClN",
  "ChemicalGroup": "Quaternary ammonium compound / antiseptic",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin/eye irritation",
  "HazardCategories": "4; 2; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H315",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=123-03-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H315, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=123-03-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C031",
  "IsNew": false,
  "ChemicalName": "Chlorhexidine digluconate",
  "Synonyms": [
   "Chlorhexidine gluconate"
  ],
  "CASNumber": "18472-51-0",
  "Formula": "C34H54Cl2N10O14",
  "ChemicalGroup": "Biguanide antiseptic",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Serious eye damage; Aquatic acute/chronic (concentrated active)",
  "HazardCategories": "1; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H318",
   "H400",
   "H410"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "0,29 (Lepomis macrochirus, 96 h; quy đổi theo chlorhexidine hoạt tính)",
   "EC50": "0,025 (Daphnia, 48 h; quy đổi); tảo 0,0062 (72 h)",
   "NOEC": "0,0116 (Daphnia, 21 ngày, NOEC); tảo EC₁₀ 0,002 (72 h)",
   "logKow": "0,08–0,5 (chlorhexidine hoạt tính)",
   "BCF": "3 ước tính; BAF quan sát ~40",
   "Biodegradation": "0% COD/28 ngày trong thử nghiệm được báo cáo → không dễ phân hủy",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm; một phần dữ liệu quy đổi theo hoạt chất",
   "EcoSource": "ECHA/Canadian screening & PubChem; CAS 18472-51-0 | https://pubchem.ncbi.nlm.nih.gov/#query=18472-51-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1; Độc tính cấp thủy sinh đáng tin cậy ≤1 mg/L (giá trị thấp nhất 0.025 mg/L); Độc tính mạn ≤0,1 mg/L và không dễ phân hủy (giá trị thấp nhất 0.002 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=18472-51-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C032",
  "IsNew": false,
  "ChemicalName": "CI 14720",
  "Synonyms": [
   "Carmoisine? / CI 14720 (cần đối chiếu nhãn)"
  ],
  "CASNumber": "4548-53-2",
  "Formula": "C20H11N2Na3O10S3",
  "ChemicalGroup": "Azo dye",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Cần đối chiếu hồ sơ định danh chính xác",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=4548-53-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=4548-53-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C033",
  "IsNew": false,
  "ChemicalName": "CI 19140",
  "Synonyms": [
   "Tartrazine",
   "FD&C Yellow No.5"
  ],
  "CASNumber": "1934-21-0",
  "Formula": "C16H9N4Na3O9S2",
  "ChemicalGroup": "Azo dye",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Không phân loại nguy hại đáng kể; có thể gây phản ứng ở người nhạy cảm",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1934-21-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1934-21-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C034",
  "IsNew": false,
  "ChemicalName": "CI 42090",
  "Synonyms": [
   "Brilliant Blue FCF",
   "FD&C Blue No.1"
  ],
  "CASNumber": "3844-45-9",
  "Formula": "C37H34N2Na2O9S3",
  "ChemicalGroup": "Synthetic dye",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể trong sử dụng thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=3844-45-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=3844-45-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C035",
  "IsNew": false,
  "ChemicalName": "Citral",
  "Synonyms": [
   "Geranial/Neral mixture"
  ],
  "CASNumber": "5392-40-5",
  "Formula": "C10H16O",
  "ChemicalGroup": "Terpenoid aldehyde",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin irritation; Skin sensitization; Eye irritation; Aquatic chronic",
  "HazardCategories": "2; 1; 2; 3",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H315",
   "H317",
   "H319",
   "H412"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có thể gây phản ứng dị ứng da.",
  "EnvEffects": "Có hại cho sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "4,1 (Oryzias latipes, 96 h; thực nghiệm)",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "3,45 (ước tính)",
   "BCF": "~87 (ước tính; có mô hình cho giá trị cao hơn)",
   "Biodegradation": "88–94% BOD/28 ngày (OECD 301C) → dễ phân hủy",
   "AquaticClass": "Aquatic Chronic 3",
   "Assessment": "Trung bình/đáng lưu ý – có phân loại Aquatic Chronic 2–4",
   "DataQuality": "Thực nghiệm + ước tính",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=5392-40-5"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H317, H319; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4; Độc tính cấp thủy sinh >1–100 mg/L (giá trị thấp nhất 4.1 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=5392-40-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C036",
  "IsNew": false,
  "ChemicalName": "Citric acid",
  "Synonyms": [
   "2-Hydroxypropane-1,2,3-tricarboxylic acid"
  ],
  "CASNumber": "77-92-9",
  "Formula": "C6H8O7",
  "ChemicalGroup": "Acid hữu cơ",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=77-92-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=77-92-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C037",
  "IsNew": false,
  "ChemicalName": "Climbazole",
  "Synonyms": [
   "1-(4-Chlorophenoxy)-1-(1H-imidazol-1-yl)-3,3-dimethylbutan-2-one"
  ],
  "CASNumber": "38083-17-9",
  "Formula": "C15H17ClN2O2",
  "ChemicalGroup": "Imidazole antifungal",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Aquatic chronic; possible skin/eye irritation",
  "HazardCategories": "2/3",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H411"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=38083-17-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=38083-17-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C038",
  "IsNew": false,
  "ChemicalName": "Clotrimazole",
  "Synonyms": [
   "Clotrimazole"
  ],
  "CASNumber": "23593-75-1",
  "Formula": "C22H17ClN2",
  "ChemicalGroup": "Imidazole antifungal",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Aquatic chronic; irritation data vary by supplier",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H411"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=23593-75-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=23593-75-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C039",
  "IsNew": false,
  "ChemicalName": "Cocamide MEA",
  "Synonyms": [
   "Coconut monoethanolamide"
  ],
  "CASNumber": "68140-00-1",
  "Formula": "Mixture/variable",
  "ChemicalGroup": "Nonionic surfactant / foam booster",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation; skin irritation possible",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=68140-00-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=68140-00-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C040",
  "IsNew": false,
  "ChemicalName": "Cocamidopropyl betaine",
  "Synonyms": [
   "CAPB"
  ],
  "CASNumber": "61789-40-0",
  "Formula": "Mixture/variable",
  "ChemicalGroup": "Amphoteric surfactant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Eye damage/irritation; skin sensitization may relate to impurities",
  "HazardCategories": "1/2",
  "Pictograms": [
   "GHS05",
   "GHS07"
  ],
  "SignalWord": "Danger/Warning",
  "HStatements": [
   "H318/H319"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=61789-40-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H318, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=61789-40-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C041",
  "IsNew": false,
  "ChemicalName": "Coumarin",
  "Synonyms": [
   "1,2-Benzopyrone"
  ],
  "CASNumber": "91-64-5",
  "Formula": "C9H6O2",
  "ChemicalGroup": "Benzopyrone / fragrance compound",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin sensitization",
  "HazardCategories": "4; 1",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H317"
  ],
  "HealthEffects": "Có thể gây phản ứng dị ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=91-64-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H317",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=91-64-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C042",
  "IsNew": false,
  "ChemicalName": "Cypermethrin",
  "Synonyms": [
   "Cypermethrin"
  ],
  "CASNumber": "52315-07-8",
  "Formula": "C22H19Cl2NO3",
  "ChemicalGroup": "Pyrethroid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; STOT SE; Aquatic acute/chronic",
  "HazardCategories": "4; 3; 1/1",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H332",
   "H335",
   "H400",
   "H410"
  ],
  "HealthEffects": "Có hại khi nuốt phải. Có hại khi hít phải. Có thể gây kích ứng đường hô hấp.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "6,60",
   "BCF": "~420–468 (quan sát ở cá)",
   "Biodegradation": "Phân hủy trong nước/đất với t½ thường khoảng 9–30 ngày; không suy diễn là ready-biodegradable",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm/field",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=52315-07-8"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=52315-07-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C043",
  "IsNew": false,
  "ChemicalName": "Deltamethrin",
  "Synonyms": [
   "Deltamethrin"
  ],
  "CASNumber": "52918-63-5",
  "Formula": "C22H19Br2NO3",
  "ChemicalGroup": "Pyrethroid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic acute/chronic",
  "HazardCategories": "3/4; 1/1",
  "Pictograms": [
   "GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301/H332",
   "H400",
   "H410"
  ],
  "HealthEffects": "Độc/nguy hiểm khi nuốt phải. Có hại khi hít phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "6,20",
   "BCF": "~417–501 (cá)",
   "Biodegradation": "Phân hủy trong bùn hoạt tính 51,4–63,4% trong 3–9 h; t½ đất thường 5–8 tuần",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm; biodegradation phụ thuộc hệ thử",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=52918-63-5"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=52918-63-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C044",
  "IsNew": false,
  "ChemicalName": "Didecyldimethylammonium chloride",
  "Synonyms": [
   "DDAC"
  ],
  "CASNumber": "7173-51-5",
  "Formula": "C22H48ClN",
  "ChemicalGroup": "Quaternary ammonium biocide",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Acute toxicity; Skin corrosion; Aquatic acute/chronic",
  "HazardCategories": "4; 1B; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H314",
   "H400",
   "H410"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "2,59 (pH 7)",
   "BCF": "81 (Lepomis macrochirus; đo)",
   "Biodegradation": "67–72%/28 ngày (OECD 301B) và 69%/28 ngày (OECD 301D) → dễ phân hủy",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm",
   "EcoSource": "PubChem/ECHA; CAS 7173-51-5 | https://pubchem.ncbi.nlm.nih.gov/#query=7173-51-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7173-51-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C045",
  "IsNew": false,
  "ChemicalName": "Diethylene glycol",
  "Synonyms": [
   "2,2'-Oxydiethanol",
   "DEG"
  ],
  "CASNumber": "111-46-6",
  "Formula": "C4H10O3",
  "ChemicalGroup": "Ether glycol / dung môi",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Acute toxicity; STOT RE",
  "HazardCategories": "4; 2",
  "Pictograms": [
   "GHS07",
   " GHS08"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H373"
  ],
  "HealthEffects": "Có hại khi nuốt phải. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=111-46-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H373",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=111-46-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C046",
  "IsNew": false,
  "ChemicalName": "Dimethoate",
  "Synonyms": [
   "Dimethoate"
  ],
  "CASNumber": "60-51-5",
  "Formula": "C5H12NO3PS2",
  "ChemicalGroup": "Organophosphate insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic acute/chronic",
  "HazardCategories": "3; 1/1",
  "Pictograms": [
   "GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H400",
   "H410"
  ],
  "HealthEffects": "Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "0,78",
   "BCF": "1,1–6 (cá; khoảng báo cáo)",
   "Biodegradation": "Biodegradation biến thiên; ~8 tuần trong nước sông, t½ đất thường 5–11 ngày",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm/field",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=60-51-5"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=60-51-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C047",
  "IsNew": false,
  "ChemicalName": "Dinotefuran",
  "Synonyms": [
   "Dinotefuran"
  ],
  "CASNumber": "165252-70-0",
  "Formula": "C7H14N4O3",
  "ChemicalGroup": "Neonicotinoid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Aquatic chronic / acute toxicity classifications vary by dossier",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=165252-70-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=165252-70-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C048",
  "IsNew": false,
  "ChemicalName": "Dipotassium glycyrrhizate",
  "Synonyms": [
   "Dipotassium glycyrrhizinate"
  ],
  "CASNumber": "68797-35-3",
  "Formula": "C42H60K2O16",
  "ChemicalGroup": "Glycyrrhizic acid salt / soothing agent",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể trong sử dụng thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=68797-35-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=68797-35-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C049",
  "IsNew": false,
  "ChemicalName": "Dipropylene glycol",
  "Synonyms": [
   "DPG"
  ],
  "CASNumber": "25265-71-8",
  "Formula": "C6H14O3",
  "ChemicalGroup": "Ether glycol / dung môi",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại theo GHS",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=25265-71-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=25265-71-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C050",
  "IsNew": false,
  "ChemicalName": "Disodium EDTA",
  "Synonyms": [
   "Disodium edetate"
  ],
  "CASNumber": "139-33-3",
  "Formula": "C10H14N2Na2O8",
  "ChemicalGroup": "Aminopolycarboxylate chelator",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation (supplier-dependent)",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=139-33-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=139-33-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C051",
  "IsNew": false,
  "ChemicalName": "DMDM hydantoin",
  "Synonyms": [
   "DMDMH"
  ],
  "CASNumber": "6440-58-0",
  "Formula": "C7H12N2O4",
  "ChemicalGroup": "Chất bảo quản giải phóng formaldehyde",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Skin sensitization; Eye irritation",
  "HazardCategories": "1; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H317",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây phản ứng dị ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=6440-58-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H317, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=6440-58-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C052",
  "IsNew": false,
  "ChemicalName": "Emamectin benzoate",
  "Synonyms": [
   "Emamectin benzoate"
  ],
  "CASNumber": "155569-91-8",
  "Formula": "C56H81NO15 (B1a benzoate; mixture with B1b)",
  "ChemicalGroup": "Avermectin insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; STOT; Aquatic acute/chronic",
  "HazardCategories": "2/3; 1; 1/1",
  "Pictograms": [
   "GHS06",
   " GHS08",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H300/H301",
   "H372",
   "H400",
   "H410"
  ],
  "HealthEffects": "Độc/nguy hiểm khi nuốt phải. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "3,0 (pH 5,1)",
   "BCF": "30 (fillet)–69 (whole fish)",
   "Biodegradation": "Phân hủy chậm trong đất: t½ ~74 ngày giai đoạn đầu, có pha kéo dài hơn",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm/field",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=155569-91-8"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H300, H301, H372; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=155569-91-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C053",
  "IsNew": false,
  "ChemicalName": "Ethanol",
  "Synonyms": [
   "Ethyl alcohol"
  ],
  "CASNumber": "64-17-5",
  "Formula": "C2H6O",
  "ChemicalGroup": "Alcohol / dung môi",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Flammable liquid; Eye irritation",
  "HazardCategories": "2; 2",
  "Pictograms": [
   "GHS02",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H225",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=64-17-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=64-17-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C054",
  "IsNew": false,
  "ChemicalName": "Ethylhexylglycerin",
  "Synonyms": [
   "Octoxyglycerin"
  ],
  "CASNumber": "70445-33-9",
  "Formula": "C11H24O3",
  "ChemicalGroup": "Glyceryl ether / preservative booster",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation (supplier-dependent)",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=70445-33-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=70445-33-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C055",
  "IsNew": false,
  "ChemicalName": "Etidronic acid",
  "Synonyms": [
   "HEDP",
   "1-Hydroxyethylidene diphosphonic acid"
  ],
  "CASNumber": "2809-21-4",
  "Formula": "C2H8O7P2",
  "ChemicalGroup": "Organophosphonate chelator",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; serious eye damage",
  "HazardCategories": "4; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=2809-21-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=2809-21-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C056",
  "IsNew": false,
  "ChemicalName": "Eucalyptol",
  "Synonyms": [
   "1,8-Cineole"
  ],
  "CASNumber": "470-82-6",
  "Formula": "C10H18O",
  "ChemicalGroup": "Monoterpene ether",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Flammable liquid; Aspiration hazard",
  "HazardCategories": "3; 1",
  "Pictograms": [
   "GHS02",
   " GHS08"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H226",
   "H304"
  ],
  "HealthEffects": "Có thể gây tử vong nếu nuốt và lọt vào đường hô hấp.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=470-82-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H304",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=470-82-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C057",
  "IsNew": false,
  "ChemicalName": "Glucose",
  "Synonyms": [
   "D-Glucose",
   "Dextrose"
  ],
  "CASNumber": "50-99-7",
  "Formula": "C6H12O6",
  "ChemicalGroup": "Monosaccharide",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=50-99-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=50-99-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C058",
  "IsNew": false,
  "ChemicalName": "Glufosinate-ammonium",
  "Synonyms": [
   "Ammonium glufosinate"
  ],
  "CASNumber": "77182-82-2",
  "Formula": "C5H18N3O4P",
  "ChemicalGroup": "Phosphinic acid herbicide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Reproductive toxicity; STOT RE",
  "HazardCategories": "4; 1B; 2",
  "Pictograms": [
   "GHS07",
   " GHS08"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H360Fd",
   "H373"
  ],
  "HealthEffects": "Có hại khi nuốt phải. Có thể gây tổn hại khả năng sinh sản hoặc thai nhi. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=77182-82-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H360",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=77182-82-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C059",
  "IsNew": false,
  "ChemicalName": "Glutathione",
  "Synonyms": [
   "Reduced glutathione",
   "GSH"
  ],
  "CASNumber": "70-18-8",
  "Formula": "C10H17N3O6S",
  "ChemicalGroup": "Tripeptide / antioxidant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=70-18-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=70-18-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C060",
  "IsNew": false,
  "ChemicalName": "Glycerol",
  "Synonyms": [
   "Glycerin",
   "Glycerine"
  ],
  "CASNumber": "56-81-5",
  "Formula": "C3H8O3",
  "ChemicalGroup": "Triol / chất giữ ẩm",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=56-81-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=56-81-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C061",
  "IsNew": false,
  "ChemicalName": "Hyaluronic acid",
  "Synonyms": [
   "Hyaluronan"
  ],
  "CASNumber": "9004-61-9",
  "Formula": "(C14H21NO11)n",
  "ChemicalGroup": "Polysaccharide / humectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=9004-61-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=9004-61-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C062",
  "IsNew": false,
  "ChemicalName": "Hydrochloric acid",
  "Synonyms": [
   "Hydrogen chloride",
   "Muriatic acid"
  ],
  "CASNumber": "7647-01-0",
  "Formula": "HCl",
  "ChemicalGroup": "Acid vô cơ mạnh",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Corrosive to metals; Skin corrosion; STOT SE",
  "HazardCategories": "1; 1B; 3",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H290",
   "H314",
   "H335"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có thể gây kích ứng đường hô hấp.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7647-01-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H314, H335",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7647-01-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C063",
  "IsNew": false,
  "ChemicalName": "Hydrogen peroxide",
  "Synonyms": [
   "Hydrogen dioxide"
  ],
  "CASNumber": "7722-84-1",
  "Formula": "H2O2",
  "ChemicalGroup": "Peroxide / chất oxy hóa",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Oxidizing liquid; Acute toxicity; Skin corrosion (concentration dependent)",
  "HazardCategories": "1; 4; 1A",
  "Pictograms": [
   "GHS03",
   " GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H271",
   "H302",
   "H332",
   "H314 (đậm đặc)"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có hại khi nuốt phải. Có hại khi hít phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7722-84-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H314, H332",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7722-84-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C064",
  "IsNew": false,
  "ChemicalName": "Imidacloprid",
  "Synonyms": [
   "Imidacloprid"
  ],
  "CASNumber": "138261-41-3",
  "Formula": "C9H10ClN5O2",
  "ChemicalGroup": "Neonicotinoid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic chronic",
  "HazardCategories": "4; 1",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H410"
  ],
  "HealthEffects": "Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=138261-41-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=138261-41-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C065",
  "IsNew": false,
  "ChemicalName": "Imiprothrin",
  "Synonyms": [
   "Imiprothrin"
  ],
  "CASNumber": "72963-72-5",
  "Formula": "C17H22N2O4",
  "ChemicalGroup": "Pyrethroid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Aquatic acute/chronic",
  "HazardCategories": "1/1",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H400",
   "H410"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "0,051 (Daphnia; EC₅₀, điều kiện thử không được mô tả đầy đủ trong tóm tắt)",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "2,9",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Dữ liệu thứ cấp trên PubChem; thiếu chi tiết thử",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=72963-72-5"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=72963-72-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C066",
  "IsNew": false,
  "ChemicalName": "Inositol",
  "Synonyms": [
   "myo-Inositol"
  ],
  "CASNumber": "87-89-8",
  "Formula": "C6H12O6",
  "ChemicalGroup": "Cyclitol",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=87-89-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=87-89-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C067",
  "IsNew": false,
  "ChemicalName": "Iodopropynyl butylcarbamate",
  "Synonyms": [
   "IPBC"
  ],
  "CASNumber": "55406-53-6",
  "Formula": "C8H12INO2",
  "ChemicalGroup": "Carbamate / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Serious eye damage; Skin sensitization; Aquatic hazard",
  "HazardCategories": "4; 1; 1; 1",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H318",
   "H317",
   "H410"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây phản ứng dị ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "36 (ước tính)",
   "Biodegradation": "Phân hủy rất nhanh trong hệ nước–trầm tích có vi sinh; t½ ~1,5 h trong điều kiện được báo cáo",
   "AquaticClass": "Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Ước tính BCF + thực nghiệm fate",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=55406-53-6"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=55406-53-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C068",
  "IsNew": false,
  "ChemicalName": "Lactic acid",
  "Synonyms": [
   "2-Hydroxypropanoic acid"
  ],
  "CASNumber": "50-21-5",
  "Formula": "C3H6O3",
  "ChemicalGroup": "Alpha-hydroxy acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin irritation; Serious eye damage",
  "HazardCategories": "2; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H315",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây kích ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=50-21-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=50-21-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C069",
  "IsNew": false,
  "ChemicalName": "Lauric acid",
  "Synonyms": [
   "Dodecanoic acid"
  ],
  "CASNumber": "143-07-7",
  "Formula": "C12H24O2",
  "ChemicalGroup": "Fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=143-07-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=143-07-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C070",
  "IsNew": false,
  "ChemicalName": "Lauryl alcohol",
  "Synonyms": [
   "1-Dodecanol"
  ],
  "CASNumber": "112-53-8",
  "Formula": "C12H26O",
  "ChemicalGroup": "Fatty alcohol",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation; Aquatic chronic",
  "HazardCategories": "2; 2",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319",
   "H411"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=112-53-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=112-53-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C071",
  "IsNew": false,
  "ChemicalName": "Linoleic acid",
  "Synonyms": [
   "9,12-Octadecadienoic acid"
  ],
  "CASNumber": "60-33-3",
  "Formula": "C18H32O2",
  "ChemicalGroup": "Polyunsaturated fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=60-33-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=60-33-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C072",
  "IsNew": false,
  "ChemicalName": "Lysozyme",
  "Synonyms": [
   "Lysozyme"
  ],
  "CASNumber": "9001-63-2",
  "Formula": "Enzyme/protein",
  "ChemicalGroup": "Muramidase",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Respiratory sensitization possible for enzyme dust",
  "HazardCategories": "1",
  "Pictograms": [
   "GHS08"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H334"
  ],
  "HealthEffects": "Có thể gây triệu chứng dị ứng hoặc hen khi hít phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=9001-63-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H334",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=9001-63-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C073",
  "IsNew": false,
  "ChemicalName": "m-Aminophenol",
  "Synonyms": [
   "3-Aminophenol"
  ],
  "CASNumber": "591-27-5",
  "Formula": "C6H7NO",
  "ChemicalGroup": "Aminophenol / hair dye intermediate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; skin/eye irritation; aquatic hazard",
  "HazardCategories": "4; 2; 2",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H315",
   "H319",
   "H411"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=591-27-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H315, H319; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=591-27-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C074",
  "IsNew": false,
  "ChemicalName": "Madecassoside",
  "Synonyms": [
   "Madecassoside"
  ],
  "CASNumber": "34540-22-2",
  "Formula": "C48H78O20",
  "ChemicalGroup": "Triterpenoid glycoside",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không có phân loại GHS hài hòa phổ biến; cần SDS nguyên liệu",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=34540-22-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=34540-22-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C075",
  "IsNew": false,
  "ChemicalName": "Mandelic acid",
  "Synonyms": [
   "2-Hydroxy-2-phenylacetic acid"
  ],
  "CASNumber": "90-64-2",
  "Formula": "C8H8O3",
  "ChemicalGroup": "Alpha-hydroxy acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=90-64-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=90-64-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C076",
  "IsNew": false,
  "ChemicalName": "Menthol",
  "Synonyms": [
   "L-Menthol"
  ],
  "CASNumber": "1490-04-6",
  "Formula": "C10H20O",
  "ChemicalGroup": "Monoterpene alcohol",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin irritation; Eye irritation",
  "HazardCategories": "2; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H315",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1490-04-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1490-04-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C077",
  "IsNew": false,
  "ChemicalName": "Mesotrione",
  "Synonyms": [
   "Mesotrione"
  ],
  "CASNumber": "104206-82-8",
  "Formula": "C14H13NO7S",
  "ChemicalGroup": "Triketone herbicide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Eye irritation; Aquatic chronic classifications may apply",
  "HazardCategories": "2; 2",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319",
   "H411"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "0,90 (pH 5); <−1 ở pH 7–9",
   "BCF": "3 (ước tính)",
   "Biodegradation": "Phân hủy là đường chính; t½ đất phòng thí nghiệm ~6–27 ngày, ngoài đồng ~3–7 ngày",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Trung bình/đáng lưu ý – có phân loại Aquatic Chronic 2–4",
   "DataQuality": "Thực nghiệm/field + ước tính BCF",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=104206-82-8"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=104206-82-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C078",
  "IsNew": false,
  "ChemicalName": "Methyl salicylate",
  "Synonyms": [
   "Wintergreen oil",
   "Methyl 2-hydroxybenzoate"
  ],
  "CASNumber": "119-36-8",
  "Formula": "C8H8O3",
  "ChemicalGroup": "Salicylate ester",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity",
  "HazardCategories": "4",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302"
  ],
  "HealthEffects": "Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=119-36-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=119-36-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C079",
  "IsNew": false,
  "ChemicalName": "Methylchloroisothiazolinone",
  "Synonyms": [
   "MCI",
   "CMIT"
  ],
  "CASNumber": "26172-55-4",
  "Formula": "C4H4ClNOS",
  "ChemicalGroup": "Isothiazolinone / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Acute toxicity; Skin corrosion; Skin sensitization; Aquatic acute/chronic",
  "HazardCategories": "2/3; 1C; 1A; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H310",
   "H314",
   "H317",
   "H400",
   "H410"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có thể gây phản ứng dị ứng da. Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=26172-55-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301, H310; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=26172-55-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C080",
  "IsNew": false,
  "ChemicalName": "Methylisothiazolinone",
  "Synonyms": [
   "MIT",
   "2-Methyl-4-isothiazolin-3-one"
  ],
  "CASNumber": "2682-20-4",
  "Formula": "C4H5NOS",
  "ChemicalGroup": "Isothiazolinone / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Acute toxicity; Skin corrosion; Skin sensitization; Aquatic acute/chronic",
  "HazardCategories": "3; 1B; 1A; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H311",
   "H314",
   "H317",
   "H400",
   "H410"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có thể gây phản ứng dị ứng da. Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "−0,486 (24 °C)",
   "BCF": "3 (Lepomis macrochirus; báo cáo)",
   "Biodegradation": "47,6–55,8% CO₂/28 ngày → inherently biodegradable; t½ ~9 ngày trong vi mô nước/sediment",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/methylisothiazolinone"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301, H311; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=2682-20-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C081",
  "IsNew": false,
  "ChemicalName": "Myristic acid",
  "Synonyms": [
   "Tetradecanoic acid"
  ],
  "CASNumber": "544-63-8",
  "Formula": "C14H28O2",
  "ChemicalGroup": "Fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=544-63-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=544-63-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C082",
  "IsNew": false,
  "ChemicalName": "Niacinamide",
  "Synonyms": [
   "Nicotinamide",
   "Vitamin B3"
  ],
  "CASNumber": "98-92-0",
  "Formula": "C6H6N2O",
  "ChemicalGroup": "Vitamin / amide",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=98-92-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=98-92-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C083",
  "IsNew": false,
  "ChemicalName": "Oleic acid",
  "Synonyms": [
   "cis-9-Octadecenoic acid"
  ],
  "CASNumber": "112-80-1",
  "Formula": "C18H34O2",
  "ChemicalGroup": "Unsaturated fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=112-80-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=112-80-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C084",
  "IsNew": false,
  "ChemicalName": "p-Phenylenediamine",
  "Synonyms": [
   "1,4-Phenylenediamine",
   "PPD"
  ],
  "CASNumber": "106-50-3",
  "Formula": "C6H8N2",
  "ChemicalGroup": "Aromatic amine / hair dye intermediate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin sensitization; Eye irritation; Aquatic hazard",
  "HazardCategories": "3/4; 1; 2; 1",
  "Pictograms": [
   "GHS06",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H312",
   "H317",
   "H319",
   "H400"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây phản ứng dị ứng da. Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=106-50-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=106-50-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C085",
  "IsNew": false,
  "ChemicalName": "Palmitic acid",
  "Synonyms": [
   "Hexadecanoic acid"
  ],
  "CASNumber": "57-10-3",
  "Formula": "C16H32O2",
  "ChemicalGroup": "Fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=57-10-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=57-10-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C086",
  "IsNew": false,
  "ChemicalName": "Panthenol",
  "Synonyms": [
   "D-Panthenol",
   "Provitamin B5"
  ],
  "CASNumber": "81-13-0",
  "Formula": "C9H19NO4",
  "ChemicalGroup": "Provitamin / humectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=81-13-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=81-13-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C087",
  "IsNew": false,
  "ChemicalName": "Papain",
  "Synonyms": [
   "Papain"
  ],
  "CASNumber": "9001-73-4",
  "Formula": "Enzyme/protein",
  "ChemicalGroup": "Proteolytic enzyme",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Respiratory sensitization; skin sensitization (enzyme dust)",
  "HazardCategories": "1; 1",
  "Pictograms": [
   "GHS08",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H334",
   "H317"
  ],
  "HealthEffects": "Có thể gây phản ứng dị ứng da. Có thể gây triệu chứng dị ứng hoặc hen khi hít phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=9001-73-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H317, H334",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=9001-73-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C088",
  "IsNew": false,
  "ChemicalName": "Patent Blue V",
  "Synonyms": [
   "Acid Blue 3",
   "E131"
  ],
  "CASNumber": "3536-49-0",
  "Formula": "C27H31N2NaO6S2",
  "ChemicalGroup": "Triarylmethane dye",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Có thể gây kích ứng/nhạy cảm ở một số người; không phân loại nghiêm trọng phổ biến",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=3536-49-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=3536-49-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C089",
  "IsNew": false,
  "ChemicalName": "Phenoxyethanol",
  "Synonyms": [
   "2-Phenoxyethanol"
  ],
  "CASNumber": "122-99-6",
  "Formula": "C8H10O2",
  "ChemicalGroup": "Ether alcohol / chất bảo quản",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Eye irritation",
  "HazardCategories": "4; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=122-99-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=122-99-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C090",
  "IsNew": false,
  "ChemicalName": "Potassium hydroxide",
  "Synonyms": [
   "Caustic potash"
  ],
  "CASNumber": "1310-58-3",
  "Formula": "KOH",
  "ChemicalGroup": "Base vô cơ mạnh",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin corrosion",
  "HazardCategories": "4; 1A",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H314"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1310-58-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H314",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1310-58-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C091",
  "IsNew": false,
  "ChemicalName": "Potassium sorbate",
  "Synonyms": [
   "Potassium 2,4-hexadienoate"
  ],
  "CASNumber": "24634-61-5",
  "Formula": "C6H7KO2",
  "ChemicalGroup": "Muối hữu cơ / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=24634-61-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=24634-61-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C092",
  "IsNew": false,
  "ChemicalName": "Profenofos",
  "Synonyms": [
   "Profenofos"
  ],
  "CASNumber": "41198-08-7",
  "Formula": "C11H15BrClO3PS",
  "ChemicalGroup": "Organophosphate insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic acute/chronic",
  "HazardCategories": "4; 1/1",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H400",
   "H410"
  ],
  "HealthEffects": "Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "4,68",
   "BCF": "29–45 (body/head), 134 (zebrafish), tới 682 (viscera)",
   "Biodegradation": "Phân hủy nhanh trong đất kiềm (t½ ~2 ngày hiếu khí) nhưng chậm hơn ở pH trung tính/axit; sản phẩm phân hủy có thể tồn tại lâu hơn",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm/field",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/profenofos"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=41198-08-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C093",
  "IsNew": false,
  "ChemicalName": "Propane",
  "Synonyms": [
   "Propyl hydride"
  ],
  "CASNumber": "74-98-6",
  "Formula": "C3H8",
  "ChemicalGroup": "Hydrocarbon / propellant",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Flammable gas; Gas under pressure",
  "HazardCategories": "1; liquefied gas",
  "Pictograms": [
   "GHS02",
   " GHS04"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H220",
   "H280"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=74-98-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=74-98-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C094",
  "IsNew": false,
  "ChemicalName": "Propanediol",
  "Synonyms": [
   "1,3-Propanediol"
  ],
  "CASNumber": "504-63-2",
  "Formula": "C3H8O2",
  "ChemicalGroup": "Diol / dung môi giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại theo GHS",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=504-63-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=504-63-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C095",
  "IsNew": false,
  "ChemicalName": "Propanol",
  "Synonyms": [
   "1-Propanol",
   "Propan-1-ol"
  ],
  "CASNumber": "71-23-8",
  "Formula": "C3H8O",
  "ChemicalGroup": "Alcohol / dung môi",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Flammable liquid; Serious eye damage; STOT SE",
  "HazardCategories": "2; 1; 3",
  "Pictograms": [
   "GHS02",
   " GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H225",
   "H318",
   "H336"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây buồn ngủ hoặc chóng mặt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=71-23-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H318, H336",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=71-23-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C096",
  "IsNew": false,
  "ChemicalName": "Propylene glycol",
  "Synonyms": [
   "1,2-Propanediol"
  ],
  "CASNumber": "57-55-6",
  "Formula": "C3H8O2",
  "ChemicalGroup": "Diol / dung môi giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại theo GHS",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=57-55-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=57-55-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C097",
  "IsNew": false,
  "ChemicalName": "Pyridoxine hydrochloride",
  "Synonyms": [
   "Vitamin B6 hydrochloride"
  ],
  "CASNumber": "58-56-0",
  "Formula": "C8H12ClNO3",
  "ChemicalGroup": "Vitamin salt",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=58-56-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=58-56-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C098",
  "IsNew": false,
  "ChemicalName": "Resorcinol",
  "Synonyms": [
   "Benzene-1,3-diol"
  ],
  "CASNumber": "108-46-3",
  "Formula": "C6H6O2",
  "ChemicalGroup": "Phenol / hair dye coupler",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin/eye irritation; Skin sensitization; Aquatic chronic",
  "HazardCategories": "4; 2; 2; 1; 3",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H315",
   "H317",
   "H319",
   "H412"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Có thể gây phản ứng dị ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Có hại cho sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "0,80",
   "BCF": "3 (ước tính)",
   "Biodegradation": "100% theoretical BOD/2 tuần (MITI); phân hủy sinh học nhanh",
   "AquaticClass": "Aquatic Chronic 3",
   "Assessment": "Trung bình/đáng lưu ý – có phân loại Aquatic Chronic 2–4",
   "DataQuality": "Thực nghiệm + ước tính BCF",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/resorcinol"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H315, H317, H319; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=108-46-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C099",
  "IsNew": false,
  "ChemicalName": "Retinyl palmitate",
  "Synonyms": [
   "Vitamin A palmitate"
  ],
  "CASNumber": "79-81-2",
  "Formula": "C36H60O2",
  "ChemicalGroup": "Retinoid ester",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Aquatic chronic; reproductive concern at high exposure requires context",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=79-81-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=79-81-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C100",
  "IsNew": false,
  "ChemicalName": "Salicylic acid",
  "Synonyms": [
   "2-Hydroxybenzoic acid"
  ],
  "CASNumber": "69-72-7",
  "Formula": "C7H6O3",
  "ChemicalGroup": "Beta-hydroxy acid / phenolic acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Serious eye damage; Reproductive toxicity",
  "HazardCategories": "4; 1; 2",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS08"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H318",
   "H361d"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có hại khi nuốt phải. Bị nghi ngờ có thể ảnh hưởng sinh sản hoặc thai nhi.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=69-72-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H318, H361",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=69-72-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C101",
  "IsNew": false,
  "ChemicalName": "Silver",
  "Synonyms": [
   "Elemental silver"
  ],
  "CASNumber": "7440-22-4",
  "Formula": "Ag",
  "ChemicalGroup": "Kim loại",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Aquatic acute/chronic",
  "HazardCategories": "1/1",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H400",
   "H410"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "~0,040 mg/L (Daphnia pulex, 48 h; nanosilver, không đại diện mọi dạng Ag)",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "Không dùng một BCF đơn nhất; Ag có thể tích lũy mô và phụ thuộc dạng/speciation",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Kim loại; endpoint phụ thuộc dạng ion/hạt nano và điều kiện nước",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/7440-22-4"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7440-22-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C102",
  "IsNew": false,
  "ChemicalName": "Sodium ascorbyl phosphate",
  "Synonyms": [
   "Trisodium ascorbyl phosphate"
  ],
  "CASNumber": "66170-10-3",
  "Formula": "C6H6Na3O9P",
  "ChemicalGroup": "Vitamin C derivative",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể; cần SDS nguyên liệu",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=66170-10-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=66170-10-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C103",
  "IsNew": false,
  "ChemicalName": "Sodium benzoate",
  "Synonyms": [
   "Benzoic acid sodium salt"
  ],
  "CASNumber": "532-32-1",
  "Formula": "C7H5NaO2",
  "ChemicalGroup": "Muối hữu cơ / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=532-32-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=532-32-1",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C104",
  "IsNew": false,
  "ChemicalName": "Sodium carbonate",
  "Synonyms": [
   "Soda ash",
   "Washing soda"
  ],
  "CASNumber": "497-19-8",
  "Formula": "Na2CO3",
  "ChemicalGroup": "Muối carbonate",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Eye irritation",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=497-19-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=497-19-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C105",
  "IsNew": false,
  "ChemicalName": "Sodium chloride",
  "Synonyms": [
   "Common salt"
  ],
  "CASNumber": "7647-14-5",
  "Formula": "NaCl",
  "ChemicalGroup": "Muối vô cơ",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân; Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại nguy hại theo GHS ở điều kiện sử dụng thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7647-14-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7647-14-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C106",
  "IsNew": false,
  "ChemicalName": "Sodium citrate",
  "Synonyms": [
   "Trisodium citrate"
  ],
  "CASNumber": "68-04-2",
  "Formula": "C6H5Na3O7",
  "ChemicalGroup": "Citrate salt / buffer",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=68-04-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=68-04-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C107",
  "IsNew": false,
  "ChemicalName": "Sodium fluoride",
  "Synonyms": [
   "Sodium fluoride"
  ],
  "CASNumber": "7681-49-4",
  "Formula": "NaF",
  "ChemicalGroup": "Muối fluoride",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin irritation; Eye irritation",
  "HazardCategories": "3; 2; 2",
  "Pictograms": [
   "GHS06",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301",
   "H315",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây kích ứng da. Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7681-49-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7681-49-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C108",
  "IsNew": false,
  "ChemicalName": "Sodium hyaluronate",
  "Synonyms": [
   "Hyaluronan sodium"
  ],
  "CASNumber": "9067-32-7",
  "Formula": "(C14H20NNaO11)n",
  "ChemicalGroup": "Polysaccharide salt / humectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=9067-32-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=9067-32-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C109",
  "IsNew": false,
  "ChemicalName": "Sodium hydroxide",
  "Synonyms": [
   "Caustic soda",
   "Lye"
  ],
  "CASNumber": "1310-73-2",
  "Formula": "NaOH",
  "ChemicalGroup": "Base vô cơ mạnh",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin corrosion",
  "HazardCategories": "1A",
  "Pictograms": [
   "GHS05"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H314"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1310-73-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H314",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1310-73-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C110",
  "IsNew": false,
  "ChemicalName": "Sodium hypochlorite",
  "Synonyms": [
   "Bleach"
  ],
  "CASNumber": "7681-52-9",
  "Formula": "NaClO",
  "ChemicalGroup": "Hypochlorite / chất oxy hóa, khử khuẩn",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Skin corrosion; Aquatic acute/chronic (concentration dependent)",
  "HazardCategories": "1B; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H314",
   "H400",
   "H410",
   "EUH031"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Chất vô cơ phản ứng mạnh; không áp dụng logKow/BCF/biodegradation theo cách của chất hữu cơ",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Sodium-hypochlorite"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7681-52-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C111",
  "IsNew": false,
  "ChemicalName": "Sodium lactate",
  "Synonyms": [
   "Sodium 2-hydroxypropanoate"
  ],
  "CASNumber": "72-17-3",
  "Formula": "C3H5NaO3",
  "ChemicalGroup": "Lactate salt / humectant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=72-17-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=72-17-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C112",
  "IsNew": false,
  "ChemicalName": "Sodium laureth sulfate",
  "Synonyms": [
   "SLES",
   "Sodium lauryl ether sulfate"
  ],
  "CASNumber": "9004-82-4",
  "Formula": "Variable ethoxylate",
  "ChemicalGroup": "Anionic surfactant",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Skin irritation; Serious eye damage (concentration dependent)",
  "HazardCategories": "2; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H315",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây kích ứng da.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=9004-82-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H315, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=9004-82-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C113",
  "IsNew": false,
  "ChemicalName": "Sodium lauryl sulfate",
  "Synonyms": [
   "Sodium dodecyl sulfate",
   "SDS/SLS"
  ],
  "CASNumber": "151-21-3",
  "Formula": "C12H25NaO4S",
  "ChemicalGroup": "Anionic surfactant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin irritation; Serious eye damage",
  "HazardCategories": "4; 2; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H315",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có thể gây kích ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=151-21-3 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H315, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=151-21-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C114",
  "IsNew": false,
  "ChemicalName": "Sodium metaperiodate",
  "Synonyms": [
   "Sodium periodate"
  ],
  "CASNumber": "7790-28-5",
  "Formula": "NaIO4",
  "ChemicalGroup": "Periodate / chất oxy hóa",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Oxidizing solid; Acute toxicity; Eye damage",
  "HazardCategories": "1; 4; 1",
  "Pictograms": [
   "GHS03",
   " GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H271",
   "H302",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7790-28-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7790-28-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C115",
  "IsNew": false,
  "ChemicalName": "Sodium methylparaben",
  "Synonyms": [
   "Sodium methyl p-hydroxybenzoate"
  ],
  "CASNumber": "5026-62-0",
  "Formula": "C8H7NaO3",
  "ChemicalGroup": "Paraben / chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Cần đối chiếu SDS; thường kích ứng nhẹ ở dạng nguyên liệu",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=5026-62-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=5026-62-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C116",
  "IsNew": false,
  "ChemicalName": "Sodium pyrithione",
  "Synonyms": [
   "Sodium omadine"
  ],
  "CASNumber": "3811-73-2",
  "Formula": "C5H4NNaOS",
  "ChemicalGroup": "Pyrithione salt / antimicrobial",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Acute toxicity; serious eye damage; aquatic hazard",
  "HazardCategories": "3/4; 1; 1",
  "Pictograms": [
   "GHS05",
   " GHS06",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H301/H302",
   "H318",
   "H400"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Độc/nguy hiểm khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=3811-73-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H301; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=3811-73-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C117",
  "IsNew": false,
  "ChemicalName": "Sodium saccharin",
  "Synonyms": [
   "Saccharin sodium"
  ],
  "CASNumber": "128-44-9",
  "Formula": "C7H4NNaO3S",
  "ChemicalGroup": "Artificial sweetener",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=128-44-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=128-44-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C118",
  "IsNew": false,
  "ChemicalName": "Sodium salicylate",
  "Synonyms": [
   "Sodium 2-hydroxybenzoate"
  ],
  "CASNumber": "54-21-7",
  "Formula": "C7H5NaO3",
  "ChemicalGroup": "Salicylate salt",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; reproductive toxicity concern",
  "HazardCategories": "4; 2",
  "Pictograms": [
   "GHS07",
   " GHS08"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H361"
  ],
  "HealthEffects": "Có hại khi nuốt phải. Bị nghi ngờ có thể ảnh hưởng sinh sản hoặc thai nhi.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=54-21-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H361",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=54-21-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C119",
  "IsNew": false,
  "ChemicalName": "Sodium sulfate",
  "Synonyms": [
   "Disodium sulfate"
  ],
  "CASNumber": "7757-82-6",
  "Formula": "Na2SO4",
  "ChemicalGroup": "Muối sulfate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7757-82-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7757-82-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C120",
  "IsNew": false,
  "ChemicalName": "Sodium sulfite",
  "Synonyms": [
   "Disodium sulfite"
  ],
  "CASNumber": "7757-83-7",
  "Formula": "Na2SO3",
  "ChemicalGroup": "Muối sulfite",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation; may release sulfur dioxide with acids",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7757-83-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7757-83-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C121",
  "IsNew": false,
  "ChemicalName": "Sorbitol",
  "Synonyms": [
   "D-Glucitol"
  ],
  "CASNumber": "50-70-4",
  "Formula": "C6H14O6",
  "ChemicalGroup": "Polyol / chất tạo ngọt, giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=50-70-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=50-70-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C122",
  "IsNew": false,
  "ChemicalName": "Stearic acid",
  "Synonyms": [
   "Octadecanoic acid"
  ],
  "CASNumber": "57-11-4",
  "Formula": "C18H36O2",
  "ChemicalGroup": "Fatty acid",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Thường không phân loại nguy hại đáng kể",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=57-11-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=57-11-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C123",
  "IsNew": false,
  "ChemicalName": "Tetrasodium EDTA",
  "Synonyms": [
   "Tetrasodium ethylenediaminetetraacetate"
  ],
  "CASNumber": "64-02-8",
  "Formula": "C10H12N2Na4O8",
  "ChemicalGroup": "Aminopolycarboxylate chelator",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; eye irritation",
  "HazardCategories": "4; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=64-02-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=64-02-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C124",
  "IsNew": false,
  "ChemicalName": "Tetrasodium pyrophosphate",
  "Synonyms": [
   "TSPP"
  ],
  "CASNumber": "7722-88-5",
  "Formula": "Na4P2O7",
  "ChemicalGroup": "Pyrophosphate salt / builder",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; serious eye damage",
  "HazardCategories": "4; 1",
  "Pictograms": [
   "GHS05",
   " GHS07"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H318"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7722-88-5 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H318",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7722-88-5",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C125",
  "IsNew": false,
  "ChemicalName": "Thiamethoxam",
  "Synonyms": [
   "Thiamethoxam"
  ],
  "CASNumber": "153719-23-4",
  "Formula": "C8H10ClN5O3S",
  "ChemicalGroup": "Neonicotinoid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Acute toxicity; Aquatic chronic",
  "HazardCategories": "4; 1",
  "Pictograms": [
   "GHS07",
   " GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H410"
  ],
  "HealthEffects": "Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "−0,13",
   "BCF": "3 (ước tính)",
   "Biodegradation": "PubChem/HSDB: chưa có dữ liệu biodegradation đất phù hợp trong hồ sơ được trích; có quang phân và thủy phân phụ thuộc pH",
   "AquaticClass": "Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Ước tính BCF; thiếu biodegradation chuẩn",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Actara"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=153719-23-4",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C126",
  "IsNew": false,
  "ChemicalName": "Thymol",
  "Synonyms": [
   "2-Isopropyl-5-methylphenol"
  ],
  "CASNumber": "89-83-8",
  "Formula": "C10H14O",
  "ChemicalGroup": "Phenolic monoterpene / antiseptic",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin corrosion; Aquatic hazard",
  "HazardCategories": "4; 1B; 2",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H314",
   "H411"
  ],
  "HealthEffects": "Gây bỏng da nghiêm trọng và tổn thương mắt. Có hại khi nuốt phải.",
  "EnvEffects": "Độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "7,8–48 (Cyprinus carpio; đo)",
   "Biodegradation": "Kết quả mâu thuẫn: MITI 0% BOD/4 tuần (không ready), nhưng các nghiên cứu khác cho phân hủy nhanh; ghi nhận ~50%/8 ngày trong nước sông",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": "Trung bình/đáng lưu ý – có phân loại Aquatic Chronic 2–4",
   "DataQuality": "Thực nghiệm; biodegradation phụ thuộc hệ thử",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Thymol"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H314; Phân loại nguy hại lâu dài thủy sinh GHS mức 2–4",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=89-83-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C127",
  "IsNew": false,
  "ChemicalName": "Titanium dioxide",
  "Synonyms": [
   "Titania"
  ],
  "CASNumber": "13463-67-7",
  "Formula": "TiO2",
  "ChemicalGroup": "Oxide kim loại / pigment",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Carcinogenicity by inhalation for certain respirable powder forms",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS08"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H351 (áp dụng cho một số dạng bột hô hấp)"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=13463-67-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H351",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=13463-67-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C128",
  "IsNew": false,
  "ChemicalName": "Tocopherol",
  "Synonyms": [
   "Vitamin E",
   "alpha-Tocopherol"
  ],
  "CASNumber": "59-02-9",
  "Formula": "C29H50O2",
  "ChemicalGroup": "Vitamin / antioxidant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=59-02-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=59-02-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C129",
  "IsNew": false,
  "ChemicalName": "Tocopheryl acetate",
  "Synonyms": [
   "alpha-Tocopheryl acetate"
  ],
  "CASNumber": "7695-91-2",
  "Formula": "C31H52O3",
  "ChemicalGroup": "Vitamin E ester / antioxidant",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7695-91-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7695-91-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C130",
  "IsNew": false,
  "ChemicalName": "Toluene-2,5-diamine sulfate",
  "Synonyms": [
   "2,5-Diaminotoluene sulfate"
  ],
  "CASNumber": "615-50-9",
  "Formula": "C7H12N2O4S",
  "ChemicalGroup": "Aromatic diamine salt / hair dye intermediate",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Skin sensitization; Eye irritation",
  "HazardCategories": "4; 1; 2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H302",
   "H317",
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt. Có thể gây phản ứng dị ứng da. Có hại khi nuốt phải.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=615-50-9 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H302, H317, H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=615-50-9",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C131",
  "IsNew": false,
  "ChemicalName": "Tranexamic acid",
  "Synonyms": [
   "trans-4-(Aminomethyl)cyclohexanecarboxylic acid"
  ],
  "CASNumber": "1197-18-8",
  "Formula": "C8H15NO2",
  "ChemicalGroup": "Amino acid derivative",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=1197-18-8 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1197-18-8",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C132",
  "IsNew": false,
  "ChemicalName": "Transfluthrin",
  "Synonyms": [
   "Transfluthrin"
  ],
  "CASNumber": "118712-89-3",
  "Formula": "C15H12Cl2F4O2",
  "ChemicalGroup": "Pyrethroid insecticide",
  "ProductCategoriesText": "Thuốc bảo vệ thực vật và hóa chất làm vườn",
  "HazardGroups": "Aquatic acute/chronic",
  "HazardCategories": "1/1",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H400",
   "H410"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Đã xác minh định danh và GHS trên PubChem; chưa trích được endpoint định lượng phù hợp",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Bayothrin"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=118712-89-3",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C133",
  "IsNew": false,
  "ChemicalName": "Triethanolamine",
  "Synonyms": [
   "TEA"
  ],
  "CASNumber": "102-71-6",
  "Formula": "C6H15NO3",
  "ChemicalGroup": "Alkanolamine / pH adjuster",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng; Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Eye irritation (classification may vary)",
  "HazardCategories": "2",
  "Pictograms": [
   "GHS07"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H319"
  ],
  "HealthEffects": "Có thể gây kích ứng mắt.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=102-71-6 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Nguy hại sức khỏe cần lưu ý theo GHS: H319",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=102-71-6",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C134",
  "IsNew": false,
  "ChemicalName": "Xylitol",
  "Synonyms": [
   "Xylite"
  ],
  "CASNumber": "87-99-0",
  "Formula": "C5H12O5",
  "ChemicalGroup": "Sugar alcohol / chất tạo ngọt",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại nguy hại theo GHS thông thường",
  "HazardCategories": "Không áp dụng",
  "Pictograms": [],
  "SignalWord": "Không áp dụng",
  "HStatements": [
   "Không áp dụng/không có H-statement ưu tiên trong phạm vi CSDL"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Chưa có phân loại nguy hại môi trường đáng kể trong phạm vi GHS tóm tắt; cần tuân thủ quy định thải bỏ.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=87-99-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=87-99-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C135",
  "IsNew": false,
  "ChemicalName": "Zinc oxide",
  "Synonyms": [
   "Zinc white"
  ],
  "CASNumber": "1314-13-2",
  "Formula": "ZnO",
  "ChemicalGroup": "Oxide kim loại",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Aquatic acute; Aquatic chronic",
  "HazardCategories": "1; 1",
  "Pictograms": [
   "GHS09"
  ],
  "SignalWord": "Warning",
  "HStatements": [
   "H400",
   "H410"
  ],
  "HealthEffects": "Không có phân loại nguy hại sức khỏe đáng kể theo GHS trong điều kiện sử dụng thông thường; vẫn cần tuân thủ nhãn/SDS.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "Không dùng BCF hữu cơ đơn nhất; Zn²⁺ có thể tích lũy sinh học phụ thuộc cân bằng nội môi/speciation",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Chất vô cơ/kim loại; fate phụ thuộc hòa tan/speciation",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/zinc_oxide"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=1314-13-2",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C136",
  "IsNew": false,
  "ChemicalName": "Zinc pyrithione",
  "Synonyms": [
   "Pyrithione zinc"
  ],
  "CASNumber": "13463-41-7",
  "Formula": "C10H8N2O2S2Zn",
  "ChemicalGroup": "Pyrithione complex / antifungal",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Reproductive toxicity; STOT RE; Aquatic acute/chronic",
  "HazardCategories": "4; 1B; 1; 1/1",
  "Pictograms": [
   "GHS07",
   " GHS08",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H360D",
   "H372",
   "H400",
   "H410"
  ],
  "HealthEffects": "Có hại khi nuốt phải. Có thể gây tổn hại khả năng sinh sản hoặc thai nhi. Có thể gây tổn thương cơ quan khi phơi nhiễm kéo dài/lặp lại.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "160–200 (Cyprinus carpio, 60 ngày; đo)",
   "Biodegradation": "t½ 2–22 h trong seawater/sediment hiếu khí; nhưng 0% BOD/4 tuần MITI ở 100 mg/L (có thể do độc với bùn hoạt tính)",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm; biodegradation phụ thuộc môi trường",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Pyrithione-Zinc"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Nguy hại sức khỏe nghiêm trọng theo GHS: H360, H372; Phân loại nguy hại thủy sinh GHS mức 1",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=13463-41-7",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C137",
  "IsNew": false,
  "ChemicalName": "Zinc sulfate",
  "Synonyms": [
   "Zinc sulphate"
  ],
  "CASNumber": "7733-02-0",
  "Formula": "ZnSO4",
  "ChemicalGroup": "Muối kẽm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Acute toxicity; Serious eye damage; Aquatic acute/chronic",
  "HazardCategories": "4; 1; 1/1",
  "Pictograms": [
   "GHS05",
   " GHS07",
   " GHS09"
  ],
  "SignalWord": "Danger",
  "HStatements": [
   "H302",
   "H318",
   "H400",
   "H410"
  ],
  "HealthEffects": "Có thể gây tổn thương mắt nghiêm trọng. Có hại khi nuốt phải.",
  "EnvEffects": "Rất độc với sinh vật thủy sinh, gây ảnh hưởng kéo dài; tránh phát thải ra môi trường.",
  "Ecotox": {
   "LC50": "11 (Anguilla japonica, 96 h); một nghiên cứu khác TLm 0,62 mg/L/96 h",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "0,145 (fathead minnow life-cycle, survival/hatchability); <0,25 (mussel embryo-larval 48 h)",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "Không áp dụng BCF hữu cơ đơn nhất; Zn²⁺ được điều hòa/tích lũy phụ thuộc sinh vật và điều kiện nước",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": "Cao – nguy hại thủy sinh GHS mức 1; ưu tiên kiểm soát phát thải",
   "DataQuality": "Thực nghiệm; chất vô cơ/ion kim loại",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/compound/Zinc-Sulfate"
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Phân loại nguy hại thủy sinh GHS mức 1; Độc tính cấp thủy sinh đáng tin cậy ≤1 mg/L (giá trị thấp nhất 0.62 mg/L)",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "PubChem (NIH/NCBI): https://pubchem.ncbi.nlm.nih.gov/#query=7733-02-0",
   "ECHA C&L Inventory: https://echa.europa.eu/regulations/clp/cl-inventory"
  ],
  "StatusNote": "Định danh hóa chất đã chuẩn hóa; GHS/H-P cần đối chiếu SDS của sản phẩm khi phân loại phụ thuộc nồng độ/dạng vật lý."
 },
 {
  "ChemicalID": "C138",
  "IsNew": false,
  "ChemicalName": "Potassium nitrite",
  "Synonyms": [
   "Potassium nitrite",
   "E249"
  ],
  "CASNumber": "7758-09-0",
  "Formula": "KNO2",
  "ChemicalGroup": "Muối nitrit / chất bảo quản thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Chất bảo quản nitrit; dữ liệu độc tính cần đánh giá theo dạng hóa chất và mức phơi nhiễm.",
  "EnvEffects": "Cần tổng hợp dữ liệu độc tính sinh thái và đặc tính môi trường từ nguồn chuẩn.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7758-09-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN (thành phần sản phẩm): https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/heo-2-lat.html",
   "PubChem: https://pubchem.ncbi.nlm.nih.gov/compound/Potassium-nitrite"
  ],
  "StatusNote": "E249; định danh hóa chất đã chuẩn hóa; cần bổ sung dữ liệu độc học/sinh thái trước khi phân HPL."
 },
 {
  "ChemicalID": "C139",
  "IsNew": false,
  "ChemicalName": "Sodium erythorbate",
  "Synonyms": [
   "Sodium erythorbate",
   "E316"
  ],
  "CASNumber": "6381-77-7",
  "Formula": "C6H7NaO6",
  "ChemicalGroup": "Muối hữu cơ / chất chống oxy hóa thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Phụ gia chống oxy hóa thực phẩm; đánh giá độc tính theo dữ liệu hóa chất chuẩn.",
  "EnvEffects": "Cần tổng hợp dữ liệu độc tính sinh thái và đặc tính môi trường từ nguồn chuẩn.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=6381-77-7 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN: https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/pate-thit-heo.html",
   "PubChem: https://pubchem.ncbi.nlm.nih.gov/compound/Sodium-erythorbate"
  ],
  "StatusNote": "E316; định danh hóa chất đã chuẩn hóa."
 },
 {
  "ChemicalID": "C140",
  "IsNew": false,
  "ChemicalName": "Sodium tripolyphosphate",
  "Synonyms": [
   "Pentasodium triphosphate",
   "STPP",
   "E451(i)"
  ],
  "CASNumber": "7758-29-4",
  "Formula": "Na5P3O10",
  "ChemicalGroup": "Phosphate / chất giữ ẩm thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Phụ gia phosphate; cần đánh giá theo dữ liệu hóa chất chuẩn và mức phơi nhiễm.",
  "EnvEffects": "Phosphate có ý nghĩa môi trường khi phát thải vào thủy vực; cần đánh giá theo dữ liệu cụ thể.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=7758-29-4 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN: https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/heo-2-lat.html",
   "PubChem: https://pubchem.ncbi.nlm.nih.gov/compound/Sodium-tripolyphosphate"
  ],
  "StatusNote": "E451(i); định danh hóa chất đã chuẩn hóa."
 },
 {
  "ChemicalID": "C141",
  "IsNew": false,
  "ChemicalName": "Sodium polyphosphate",
  "Synonyms": [
   "Sodium polyphosphates",
   "E452(i)"
  ],
  "CASNumber": "68915-31-1",
  "Formula": "(NaPO3)n",
  "ChemicalGroup": "Polyphosphate / chất giữ ẩm thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Nhóm polyphosphate dùng làm phụ gia; đặc tính phụ thuộc thành phần/độ trùng hợp.",
  "EnvEffects": "Cần đánh giá tác động môi trường theo dạng polyphosphate cụ thể.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "BCF": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "Biodegradation": "N/A – không áp dụng/không có ý nghĩa cho dạng hóa chất này",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=68915-31-1 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN: https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/pate-thit-heo.html"
  ],
  "StatusNote": "E452(i); dữ liệu có tính nhóm, cần thận trọng khi gán CAS/GHS."
 },
 {
  "ChemicalID": "C142",
  "IsNew": false,
  "ChemicalName": "Monosodium glutamate",
  "Synonyms": [
   "Sodium glutamate",
   "MSG",
   "E621"
  ],
  "CASNumber": "142-47-2",
  "Formula": "C5H8NNaO4",
  "ChemicalGroup": "Muối amino acid / chất điều vị thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Chất điều vị thực phẩm; đánh giá theo dữ liệu hóa chất chuẩn.",
  "EnvEffects": "Cần tổng hợp dữ liệu độc tính sinh thái và đặc tính môi trường từ nguồn chuẩn.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=142-47-2 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN: https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/heo-2-lat.html",
   "PubChem: https://pubchem.ncbi.nlm.nih.gov/compound/Monosodium-glutamate"
  ],
  "StatusNote": "E621; định danh hóa chất đã chuẩn hóa."
 },
 {
  "ChemicalID": "C143",
  "IsNew": false,
  "ChemicalName": "Erythrosine",
  "Synonyms": [
   "Erythrosine B",
   "FD&C Red No. 3",
   "E127"
  ],
  "CASNumber": "16423-68-0",
  "Formula": "C20H6I4Na2O5",
  "ChemicalGroup": "Phẩm màu xanthene / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Cần đối chiếu GHS/SDS theo nguồn hóa chất",
  "HazardCategories": "Cần đối chiếu",
  "Pictograms": [],
  "SignalWord": "Cần đối chiếu",
  "HStatements": [
   "Cần đối chiếu H statements theo SDS"
  ],
  "HealthEffects": "Phẩm màu thực phẩm; cần đánh giá độc tính theo dữ liệu hóa chất và mức phơi nhiễm.",
  "EnvEffects": "Cần tổng hợp dữ liệu độc tính sinh thái và đặc tính môi trường từ nguồn chuẩn.",
  "Ecotox": {
   "LC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "EC50": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "NOEC": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "logKow": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "BCF": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "Biodegradation": "NF – chưa tìm thấy dữ liệu định lượng phù hợp trong nguồn ưu tiên",
   "AquaticClass": "Chưa có phân loại Aquatic trong H-statements hiện có",
   "Assessment": "Chưa đủ dữ liệu định lượng – đánh giá sơ bộ dựa trên GHS/H-statements hiện có",
   "DataQuality": "Tra cứu theo CAS; chưa tìm thấy endpoint định lượng chuẩn hóa phù hợp trong nguồn ưu tiên",
   "EcoSource": "https://pubchem.ncbi.nlm.nih.gov/#query=16423-68-0 | https://echa.europa.eu/information-on-chemicals | https://www.echemportal.org/echemportal/"
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Không ghi nhận tiêu chí HPL-2/HPL-3 từ bộ dữ liệu được lựa chọn. HPL-1 chỉ phản ánh mức ưu tiên cảnh báo thấp trên dữ liệu hiện có, không đồng nghĩa hóa chất an toàn.",
  "HPLVerification": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "HPLSourceURL": null,
  "CheckNotes": null,
  "References": [
   "VISSAN: https://www.vissan.com.vn/san-pham/thuc-pham-che-bien/do-hop/pate-thit-heo.html",
   "PubChem: https://pubchem.ncbi.nlm.nih.gov/compound/Erythrosine"
  ],
  "StatusNote": "E127; định danh hóa chất đã chuẩn hóa."
 },
 {
  "ChemicalID": "C144",
  "IsNew": true,
  "ChemicalName": "2,2'-Methylenebis(6-tert-butyl-4-methylphenol)",
  "Synonyms": [],
  "CASNumber": "119-47-1",
  "Formula": null,
  "ChemicalGroup": "Chất chống oxy hóa",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H360F"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Độc tính sinh sản, nhóm 1B; có thể gây ảnh hưởng khả năng sinh sản.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://hcis.safeworkaustralia.gov.au/HazardousChemical/Details?chemicalID=5514",
  "CheckNotes": "H360F áp dụng cho chất CAS 119-47-1, không suy ra phân loại hỗn hợp keo.",
  "References": [
   "https://www.scribd.com/document/939226699/Selleys-Liquid-Nails-Original-nz-Sds"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C145",
  "IsNew": true,
  "ChemicalName": "2-[2-(2-Butoxyethoxy)ethoxy]ethanol",
  "Synonyms": [],
  "CASNumber": "143-22-6",
  "Formula": null,
  "ChemicalGroup": "Glycol ether / dung môi",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H318"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Tổn thương mắt nghiêm trọng, nhóm 1.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://hcis.safeworkaustralia.gov.au/HazardousChemical/Details?chemicalID=696",
  "CheckNotes": "HPL của hóa chất; không suy ra rủi ro sản phẩm.",
  "References": [
   "https://msdspds.castrol.com/ussds/amersdsf.nsf/Files/51C1B693F38965B580258C7C00529896/%24File/3180068.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C146",
  "IsNew": true,
  "ChemicalName": "Acetylated distarch adipate (E1422)",
  "Synonyms": [],
  "CASNumber": "68130-14-3 (cần xác minh loại tinh bột)",
  "Formula": null,
  "ChemicalGroup": "Tinh bột biến tính / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C147",
  "IsNew": true,
  "ChemicalName": "Agar-agar",
  "Synonyms": [],
  "CASNumber": "9002-18-0",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: SDS của agar (CAS 9002-18-0) không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.neogen.com/49be68/globalassets/pim/assets/original/10048/ncm0205_700003497_sds_en-us.pdf",
  "CheckNotes": "Chỉ trên cơ sở SDS được lựa chọn; bụi vẫn có thể gây kích ứng.",
  "References": [
   "https://rovinco.vn/cua-hang/bot-rau-cau-gion-rovin-25g/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C148",
  "IsNew": true,
  "ChemicalName": "Ammonium lauryl sulfate",
  "Synonyms": [],
  "CASNumber": "2235-54-3",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt anion",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H318",
   "H319"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Kích ứng da/tổn thương hoặc kích ứng mắt theo các thông báo C&L.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://pubchem.ncbi.nlm.nih.gov/compound/Ammonium-Lauryl-Sulfate",
  "CheckNotes": "Các thông báo phân loại không hoàn toàn đồng nhất; cần SDS của nguyên liệu cụ thể.",
  "References": [
   "https://finecare.com.vn/products/dau-goi-phuc-hoi-da-tang-nam-truffle-weilaiya-white-truffle-lightness-shampoo-450ml"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C149",
  "IsNew": true,
  "ChemicalName": "BHT",
  "Synonyms": [],
  "CASNumber": "128-37-0",
  "Formula": null,
  "ChemicalGroup": "Chất chống oxy hóa",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H400",
   "H410"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Aquatic Acute 1 và Aquatic Chronic 1: rất độc với sinh vật thủy sinh.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.carlroth.com/medias/SDB-NC05-GB-EN.pdf?context=bWFzdGVyfHNlY3VyaXR5RGF0YXNoZWV0c3wyODI3NjV8YXBwbGljYXRpb24vcGRmfGFEQmxMMmhsTXk4NU1qSXpNVGN4TnpNMU5UZ3lMMU5FUWw5T1F6QTFYMGRDWDBWT0xuQmtaZ3wyN2EwYjgzZDY1YTcxMGU0NTU4NGJlMTVhNGI1YmVjMTNiZWIyYmRiZGYwMjIzNWQzMWZlOTJkYWE3ZTI3ZDM3",
  "CheckNotes": "BHT là phụ gia/chất chống oxy hóa; HPL không đánh giá rủi ro theo hàm lượng.",
  "References": [
   "https://cooponline.vn/sua-tam-lifebuoy-matcha-kho-qua-sach-sau-800g--s250109901; https://www.sunsilk.com.vn/p/dau-goi-ong-muot-rang-ngoi.html/08934868156925"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C150",
  "IsNew": true,
  "ChemicalName": "Benzalkonium chloride (C12–C16 alkylbenzyl dimethyl ammonium chlorides)",
  "Synonyms": [],
  "CASNumber": "68424-85-1 (hỗn hợp homolog)",
  "Formula": null,
  "ChemicalGroup": "Hợp chất amoni bậc bốn / chất kháng khuẩn",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H302",
   "H312",
   "H314",
   "H318",
   "H400",
   "H410"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Amoni bậc bốn; ăn mòn và Aquatic Acute/Chronic 1.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://file.hstatic.net/200000985023/file/ms4800132_vn02_vi_spectrum.pdf",
  "CheckNotes": "Hỗn hợp đồng đẳng C12–C16; đối chiếu tỷ lệ đồng đẳng của nguyên liệu.",
  "References": [
   "https://cdn1.npcdn.net/userfiles/25756/attachment/5059085/20240625_22832d4380eb63bba6c81c09af61e86d/3-IN-ONE-Air-Conditioner-Cleaner-ASIA-GHS-SDS.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C151",
  "IsNew": true,
  "ChemicalName": "Benzisothiazolinone",
  "Synonyms": [],
  "CASNumber": "2634-33-5",
  "Formula": null,
  "ChemicalGroup": "Chất bảo quản",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H330",
   "H302",
   "H315",
   "H318",
   "H317",
   "H400",
   "H410"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: BIT: độc tính hít cấp nghiêm trọng và Aquatic Acute/Chronic 1 theo CLP cập nhật.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://recherche.chemikalieninfo.de/public/stoff/6032",
  "CheckNotes": "Cần áp dụng phiên bản CLP hiện hành, không dùng phân loại đề xuất cũ.",
  "References": [
   "https://lixco.com/san-pham/lix/rua-chen/sieu-dam-dac/nuoc-rua-chen-lix-sieu-dam-dac-huong-tra-xanh-172.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C152",
  "IsNew": true,
  "ChemicalName": "Benzophenone-4",
  "Synonyms": [],
  "CASNumber": "4065-45-6",
  "Formula": null,
  "ChemicalGroup": "Chất hấp thụ UV",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H317",
   "H319"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Kích ứng da/mắt và mẫn cảm da theo tổng hợp thông báo ECHA C&L.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://pubchem.ncbi.nlm.nih.gov/compound/benzophenone-4",
  "CheckNotes": "Phân loại thông báo, không phải toàn bộ nhà cung cấp đều thống nhất.",
  "References": [
   "https://cocolux.com/sua-tam-lux-botanicals-huong-hoa-thien-dieu-570g-i.8934868187875"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C153",
  "IsNew": true,
  "ChemicalName": "Beta-carotene (E160a(i))",
  "Synonyms": [],
  "CASNumber": "7235-40-7",
  "Formula": null,
  "ChemicalGroup": "Carotenoid / màu thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: SDS β-carotene CAS 7235-40-7 không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.merckmillipore.com/Web-CH-Site/en_US/-/CHF/ShowDocument-File?Country=MY&DocumentId=217538_SDS_MY_EN.PDF&DocumentType=MSD&DocumentUID=354949000&Language=EN&Origin=null&ProductSKU=EMD_BIO-217538",
  "CheckNotes": "Không đánh đồng HPL-1 với an toàn tuyệt đối ở mọi liều lượng.",
  "References": [
   "https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C154",
  "IsNew": true,
  "ChemicalName": "Betaine",
  "Synonyms": [],
  "CASNumber": "107-43-7",
  "Formula": null,
  "ChemicalGroup": "Chất giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://www.canova.it/en/mild-and-moderate-acne/287-acnell-crema-restitutiva.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C155",
  "IsNew": true,
  "ChemicalName": "Carbon dioxide",
  "Synonyms": [],
  "CASNumber": "124-38-9",
  "Formula": null,
  "ChemicalGroup": "Khí / chất đẩy",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Khí chịu áp suất (nguy hại vật lý)",
  "HazardCategories": "Press. Gas; không đưa vào điểm HPL",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H280 (nguy hại vật lý)"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: CO₂ được phân loại khí chịu áp suất; chưa có tiêu chí sức khỏe/sinh thái HPL-2/3 trong SDS được lựa chọn.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://sg.airliquide.com/sites/al_sg/files/2020/08/06/2010481_carbon_dioxide124-38-9_r2.pdf",
  "CheckNotes": "Ngạt trong không gian kín và nguy hại áp suất phải cảnh báo RIÊNG, không quy đổi trực tiếp sang HPL.",
  "References": [
   "https://media.wd40.asia/app/uploads/2024/01/26011053/WD-40-Specialist-Contact-Cleaner-Asia-GHS-SDS-15-11-23.pdf; https://media.wd40.asia/app/uploads/2024/02/26010910/SDS_WD-40-Multi-Use-Product-Aerosol_VIET.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C156",
  "IsNew": true,
  "ChemicalName": "Carrageenan",
  "Synonyms": [],
  "CASNumber": "9000-07-1",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://phacheviet.com/products/bot-rau-cau-deo-rovin-hoang-yen-hop-10-goi-120g"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C157",
  "IsNew": true,
  "ChemicalName": "Cocamide DEA",
  "Synonyms": [],
  "CASNumber": "68603-42-9 (hỗn hợp)",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H318",
   "H411"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Cocamide DEA: kích ứng da, tổn thương mắt và Aquatic Chronic 2.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://baza.easybase.pl/media/img/original/14/04/2/1404298.pdf",
  "CheckNotes": "SDS ghi CAS 68603-42-9; đối chiếu thành phần hỗn hợp nguyên liệu.",
  "References": [
   "https://icheck.vn/san-pham/sua-rua-mat-acnevir-santafa-do-8938535677218"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C158",
  "IsNew": true,
  "ChemicalName": "Coco-glucoside",
  "Synonyms": [],
  "CASNumber": "141464-42-8 (hỗn hợp alkyl polyglucoside)",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt không ion",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H318",
   "H319"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Coco-glucoside: có thông báo GHS về kích ứng da/mắt.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://storage.googleapis.com/safety-data-sheets/Manufacturer%27s%20Declarations%20-%20Hygiene%20AU/Betasan%20Max%20-%20Hand%2C%20Hair%20%26%20Body%20Wash",
  "CheckNotes": "Hỗn hợp alkyl glucoside; phân loại phụ thuộc phân bố mạch và SDS nguyên liệu.",
  "References": [
   "https://finecare.com.vn/products/dau-goi-phuc-hoi-da-tang-nam-truffle-weilaiya-white-truffle-lightness-shampoo-450ml"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C159",
  "IsNew": true,
  "ChemicalName": "Deceth-8",
  "Synonyms": [],
  "CASNumber": "26183-52-8",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt không ion",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H302",
   "H315",
   "H318"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Deceth ethoxylate CAS 26183-52-8: độc tính uống mức 4, kích ứng da, tổn thương mắt.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://hcis.safeworkaustralia.gov.au/HazardousChemical/Details?chemicalID=6506",
  "CheckNotes": "CAS có thể bao gồm số mắt xích EO khác nhau; cần đối chiếu Deceth-8 chính xác.",
  "References": [
   "https://assets.office-discount.at/img/28/11/i1356683.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C160",
  "IsNew": true,
  "ChemicalName": "Dicalcium phosphate",
  "Synonyms": [],
  "CASNumber": "7757-93-9",
  "Formula": null,
  "ChemicalGroup": "Phụ gia thực phẩm E341(ii)",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C161",
  "IsNew": true,
  "ChemicalName": "Diisopropanolamine",
  "Synonyms": [],
  "CASNumber": "110-97-4",
  "Formula": null,
  "ChemicalGroup": "Amino alcohol",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H319"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Diisopropanolamine: gây kích ứng mắt theo SDS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.actylislab.com/msds/Diisopropanolamine.pdf",
  "CheckNotes": "HPL của hóa chất; không suy ra rủi ro sản phẩm.",
  "References": [
   "https://msdspds.castrol.com/ussds/amersdsf.nsf/Files/51C1B693F38965B580258C7C00529896/%24File/3180068.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C162",
  "IsNew": true,
  "ChemicalName": "Dimethicone",
  "Synonyms": [],
  "CASNumber": "9006-65-9 (polymer)",
  "Formula": null,
  "ChemicalGroup": "Silicone / chất làm mềm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H319"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Có nguồn nhưng cần xác minh dạng chất",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Dimethicone: có thông báo C&L H319; các nhà cung cấp khác báo không phân loại.",
  "HPLVerification": "CẦN XÁC MINH",
  "HPLSourceURL": "https://pubchem.ncbi.nlm.nih.gov/compound/9006-65-9",
  "CheckNotes": "HPL-2 là mức sàng lọc bảo thủ; cần kiểm tra grade, độ nhớt và SDS đúng dạng polymer.",
  "References": [
   "https://www.canova.it/en/mild-and-moderate-acne/287-acnell-crema-restitutiva.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C163",
  "IsNew": true,
  "ChemicalName": "Dipotassium hydrogen phosphate (E340(ii))",
  "Synonyms": [],
  "CASNumber": "7758-11-4",
  "Formula": null,
  "ChemicalGroup": "Muối phosphat / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C164",
  "IsNew": true,
  "ChemicalName": "Disodium diphosphate (E450(i))",
  "Synonyms": [],
  "CASNumber": "7758-16-9 (dạng thường gặp; cần đối chiếu nhãn)",
  "Formula": null,
  "ChemicalGroup": "Phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C165",
  "IsNew": true,
  "ChemicalName": "Disodium hydrogen phosphate (E339(ii))",
  "Synonyms": [],
  "CASNumber": "7558-79-4",
  "Formula": null,
  "ChemicalGroup": "Muối phosphat / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C166",
  "IsNew": true,
  "ChemicalName": "Formic acid",
  "Synonyms": [],
  "CASNumber": "64-18-6",
  "Formula": null,
  "ChemicalGroup": "Axit hữu cơ",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H314"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Axit formic: ăn mòn da và gây tổn thương mắt nghiêm trọng.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://pubchem.ncbi.nlm.nih.gov/compound/formic_acid",
  "CheckNotes": "Phân loại chất tinh khiết; nồng độ dung dịch có thể thay đổi nguy hại.",
  "References": [
   "https://assets.office-discount.at/img/28/11/i1356683.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C167",
  "IsNew": true,
  "ChemicalName": "Gellan gum (E418)",
  "Synonyms": [],
  "CASNumber": "71010-52-1",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C168",
  "IsNew": true,
  "ChemicalName": "Glucono delta-lactone",
  "Synonyms": [],
  "CASNumber": "90-80-2",
  "Formula": null,
  "ChemicalGroup": "Phụ gia thực phẩm E575",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C169",
  "IsNew": true,
  "ChemicalName": "Glycolic acid",
  "Synonyms": [],
  "CASNumber": "79-14-1",
  "Formula": null,
  "ChemicalGroup": "Axit hydroxycarboxylic",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H332",
   "H314",
   "H318"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Axit glycolic: có hại khi hít phải, ăn mòn da và tổn thương mắt.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.merckmillipore.com/Web-JP-Site/ja_JP/-/JPY/ShowDocument-File?Country=VN&DocumentId=804104_SDS_VN_VI.PDF&DocumentType=MSD&DocumentUID=3619297&Language=VI&Origin=null&ProductSKU=MDA_CHEM-804104",
  "CheckNotes": "HPL của chất; nguy hại sản phẩm còn phụ thuộc nồng độ và pH.",
  "References": [
   "https://icheck.vn/san-pham/sua-rua-mat-acnevir-santafa-do-8938535677218"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C170",
  "IsNew": true,
  "ChemicalName": "Guar gum (E412)",
  "Synonyms": [],
  "CASNumber": "9000-30-0",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Guar gum CAS 9000-30-0: SDS không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.guinama.com/documentacion-tecnica/93249_FDS-Guar-gum-EN-Laboratorios-GUINAMA.pdf",
  "CheckNotes": "Nguy hại bụi cháy/nổ và bụi hô hấp cần xem riêng.",
  "References": [
   "https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C171",
  "IsNew": true,
  "ChemicalName": "Hexane isomers (mixture)",
  "Synonyms": [],
  "CASNumber": "92112-69-1",
  "Formula": null,
  "ChemicalGroup": "Hỗn hợp hydrocarbon",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://media.wd40.asia/app/uploads/2024/01/26011053/WD-40-Specialist-Contact-Cleaner-Asia-GHS-SDS-15-11-23.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C172",
  "IsNew": true,
  "ChemicalName": "Hydroxyethyl cellulose",
  "Synonyms": [],
  "CASNumber": "9004-62-0",
  "Formula": null,
  "ChemicalGroup": "Polymer / chất tạo đặc",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://lixco.com/san-pham/lix/rua-chen/sieu-dam-dac/nuoc-rua-chen-lix-sieu-dam-dac-huong-tra-xanh-172.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C173",
  "IsNew": true,
  "ChemicalName": "Hydroxypropyl methylcellulose",
  "Synonyms": [],
  "CASNumber": "9004-65-3",
  "Formula": null,
  "ChemicalGroup": "Polymer / chất tạo đặc",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://cooponline.vn/sua-tam-lifebuoy-matcha-kho-qua-sach-sau-800g--s250109901; https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C174",
  "IsNew": true,
  "ChemicalName": "Konjac glucomannan",
  "Synonyms": [],
  "CASNumber": "37220-17-0",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://phacheviet.com/products/bot-rau-cau-deo-rovin-hoang-yen-hop-10-goi-120g"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C175",
  "IsNew": true,
  "ChemicalName": "Lactic acid esters of mono-/diglycerides (E472b)",
  "Synonyms": [],
  "CASNumber": "CAS chưa xác định (hỗn hợp)",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html; https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C176",
  "IsNew": true,
  "ChemicalName": "Lecithins (E322(i))",
  "Synonyms": [],
  "CASNumber": "8002-43-5 (hỗn hợp lecithin)",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C177",
  "IsNew": true,
  "ChemicalName": "Magnesium sulfate",
  "Synonyms": [],
  "CASNumber": "7487-88-9 (dạng khan; hydrat chưa xác định)",
  "Formula": null,
  "ChemicalGroup": "Muối vô cơ",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Magnesium sulfate khan CAS 7487-88-9: SDS không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.sigmaaldrich.com/US/en/sds/sigald/m7506?srsltid=AfmBOopfQ0UjFSVz7PdZ7jJUwZhB7ebokeuy4u_W77G8NiJn9MtAhHAV",
  "CheckNotes": "Cần xác minh dạng khan hay ngậm nước trong sản phẩm.",
  "References": [
   "https://lixco.com/san-pham/lix/rua-chen/sieu-dam-dac/nuoc-rua-chen-lix-sieu-dam-dac-huong-tra-xanh-172.html; https://www.bachhoaxanh.com/nuoc-rua-chen/nuoc-rua-chen-gift-huong-chanh-can-4kg/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C178",
  "IsNew": true,
  "ChemicalName": "Maltodextrin",
  "Synonyms": [],
  "CASNumber": "9050-36-6",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Maltodextrin CAS 9050-36-6: SDS không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.carlroth.com/medias/SDB-2260-IE-EN.pdf?context=bWFzdGVyfHNlY3VyaXR5RGF0YXNoZWV0c3wyMzc0NTl8YXBwbGljYXRpb24vcGRmfGFEYzFMMmcwWlM4NU1UZ3pORGd4TmpVNU5ESXlMMU5FUWw4eU1qWXdYMGxGWDBWT0xuQmtaZ3xlZDQyMDIyOWU0MzNiNWYzZDdiNTc5NGRhOGJjYzVjYWEyNmU5NmE3YTUwMjA3YzllYTI4MzE4ZmY2NzQ0NzFl",
  "CheckNotes": "Bụi cháy nổ là nguy hại vật lý cần xem riêng.",
  "References": [
   "https://finecare.com.vn/products/dau-goi-phuc-hoi-da-tang-nam-truffle-weilaiya-white-truffle-lightness-shampoo-450ml"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C179",
  "IsNew": true,
  "ChemicalName": "Methylcellulose (E461)",
  "Synonyms": [],
  "CASNumber": "9004-67-5",
  "Formula": null,
  "ChemicalGroup": "Polymer / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C180",
  "IsNew": true,
  "ChemicalName": "Mono- and diglycerides of fatty acids (E471)",
  "Synonyms": [],
  "CASNumber": "CAS chưa xác định (hỗn hợp)",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C181",
  "IsNew": true,
  "ChemicalName": "Monocalcium phosphate",
  "Synonyms": [],
  "CASNumber": "7758-23-8",
  "Formula": null,
  "ChemicalGroup": "Phụ gia thực phẩm E341(i)",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C182",
  "IsNew": true,
  "ChemicalName": "Naphtha (petroleum), hydrotreated light",
  "Synonyms": [],
  "CASNumber": "64742-49-0",
  "Formula": null,
  "ChemicalGroup": "Hỗn hợp hydrocarbon dầu mỏ",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H304",
   "H315",
   "H336",
   "H411"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": null,
   "DataQuality": "Có nguồn nhưng cần xác minh dạng chất",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Naphtha hydrotreated light: nguy hại hít sặc, kích ứng da, ảnh hưởng thần kinh cấp và Aquatic Chronic 2.",
  "HPLVerification": "CẦN XÁC MINH",
  "HPLSourceURL": "https://assets.thermofisher.com/directwebviewer/private/results.aspx?LANGUAGE=d__EE&PLANT=d__ACR&SKU=ACR39692&SUBFORMAT=d__CLP1&page=NewSearch",
  "CheckNotes": "UVCB dầu mỏ: phân loại Carc./Muta. có thể phụ thuộc Note P/hàm lượng benzene; cần SDS đúng nguyên liệu để chốt HPL.",
  "References": [
   "https://www.scribd.com/document/939226699/Selleys-Liquid-Nails-Original-nz-Sds"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C183",
  "IsNew": true,
  "ChemicalName": "Petroleum distillates, hydrotreated light",
  "Synonyms": [],
  "CASNumber": "64742-47-8",
  "Formula": null,
  "ChemicalGroup": "Hỗn hợp hydrocarbon dầu mỏ",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H304"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Petroleum distillates hydrotreated light: nguy hại hít sặc (Asp. Tox. 1).",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://hcis.safeworkaustralia.gov.au/HazardousChemical/Details?chemicalID=1717",
  "CheckNotes": "Dạng hỗn hợp hydrocarbon C9–C16; nguy hại cháy cần cảnh báo riêng.",
  "References": [
   "https://media.wd40.asia/app/uploads/2024/02/26010910/SDS_WD-40-Multi-Use-Product-Aerosol_VIET.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C184",
  "IsNew": true,
  "ChemicalName": "Piroctone olamine",
  "Synonyms": [],
  "CASNumber": "68890-66-4",
  "Formula": null,
  "ChemicalGroup": "Chất kháng nấm / bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H318",
   "H400",
   "H410"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Acute 1; Aquatic Chronic 1",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 3,
  "HPLBasis": "HPL-3: Piroctone olamine: Aquatic Acute 1 và Aquatic Chronic 1.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://pubchem.ncbi.nlm.nih.gov/compound/piroctone-olamine",
  "CheckNotes": "Có khác biệt giữa thông báo C&L; H400/H410 được ghi nhận trong nguồn.",
  "References": [
   "https://cooponline.vn/sua-tam-lifebuoy-matcha-kho-qua-sach-sau-800g--s250109901; https://upharma.vn/thong-tin-san-pham/dau-goi-duoc-lieu-nguyen-xuan-vang-huong-hoa-bong-benh-cho-toc-%28hop-1-chai-250ml%29-mp02478; https://www.pharmacity.vn/dau-goi-clear-9-thao-duoc-co-truyen-lam-sach-gau-350g.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C185",
  "IsNew": true,
  "ChemicalName": "Polyglycerol esters of fatty acids (E475)",
  "Synonyms": [],
  "CASNumber": "CAS chưa xác định (hỗn hợp)",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C186",
  "IsNew": true,
  "ChemicalName": "Polysorbate 60 (E435)",
  "Synonyms": [],
  "CASNumber": "9005-67-8",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C187",
  "IsNew": true,
  "ChemicalName": "Potassium azeloyl diglycinate",
  "Synonyms": [],
  "CASNumber": "CAS cần xác minh",
  "Formula": null,
  "ChemicalGroup": "Hoạt chất mỹ phẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://www.canova.it/en/mild-and-moderate-acne/287-acnell-crema-restitutiva.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C188",
  "IsNew": true,
  "ChemicalName": "Potassium chloride",
  "Synonyms": [],
  "CASNumber": "7447-40-7",
  "Formula": null,
  "ChemicalGroup": "Muối vô cơ",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Potassium chloride CAS 7447-40-7: SDS không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.selleckchem.com/products/potassium-chloride/msds.html",
  "CheckNotes": "HPL của hóa chất; không suy ra rủi ro sản phẩm.",
  "References": [
   "https://phacheviet.com/products/bot-rau-cau-deo-rovin-hoang-yen-hop-10-goi-120g"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C189",
  "IsNew": true,
  "ChemicalName": "Propylene glycol esters of fatty acids (E477)",
  "Synonyms": [],
  "CASNumber": "CAS chưa xác định (hỗn hợp)",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C190",
  "IsNew": true,
  "ChemicalName": "Sodium PCA",
  "Synonyms": [],
  "CASNumber": "28874-51-3",
  "Formula": null,
  "ChemicalGroup": "Chất giữ ẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://icheck.vn/san-pham/sua-rua-mat-acnevir-santafa-do-8938535677218"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C191",
  "IsNew": true,
  "ChemicalName": "Sodium alkylbenzene sulfonate",
  "Synonyms": [],
  "CASNumber": "CAS chưa xác định (hỗn hợp alkylbenzene sulfonate)",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt anion",
  "ProductCategoriesText": "Chất tẩy rửa và vệ sinh gia dụng",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://lixco.com/san-pham/lix/rua-chen/sieu-dam-dac/nuoc-rua-chen-lix-sieu-dam-dac-huong-tra-xanh-172.html; https://www.bachhoaxanh.com/nuoc-rua-chen/nuoc-rua-chen-gift-huong-chanh-can-4kg/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C192",
  "IsNew": true,
  "ChemicalName": "Sodium bicarbonate",
  "Synonyms": [],
  "CASNumber": "144-55-8",
  "Formula": null,
  "ChemicalGroup": "Muối vô cơ / phụ gia thực phẩm E500(ii)",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Sodium bicarbonate CAS 144-55-8: SDS không phân loại nguy hại GHS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.merckmillipore.com/Web-INTL-Site/en_US/-/USD/ShowDocument-File?Country=SG&DocumentId=172577_SDS_SG_EN.PDF&DocumentType=MSD&DocumentUID=356570836&Language=EN&Origin=null&ProductSKU=MDA_CHEM-172577",
  "CheckNotes": "HPL của hóa chất; không suy ra rủi ro sản phẩm.",
  "References": [
   "https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C193",
  "IsNew": true,
  "ChemicalName": "Sodium stearoyl lactylate (E481(i))",
  "Synonyms": [],
  "CASNumber": "25383-99-7",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://vinmart.co/bot-banh-muffin-richmix-vi-vani-tui-1kg-l1932297432.html; https://www.nguyenlieutanphu.com/kem-trang-tri-banh-richs-nhan-vang-gold-label; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C194",
  "IsNew": true,
  "ChemicalName": "Sorbic acid",
  "Synonyms": [],
  "CASNumber": "110-44-1",
  "Formula": null,
  "ChemicalGroup": "Chất bảo quản",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H315",
   "H319",
   "H335"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Sorbic acid: kích ứng da, mắt và đường hô hấp theo SDS.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://actylislab.com/msds/Sorbic%20acid.pdf",
  "CheckNotes": "Phân loại chất nguyên liệu không thay thế đánh giá liều trong thực phẩm.",
  "References": [
   "https://www.sunsilk.com.vn/p/dau-goi-ong-muot-rang-ngoi.html/08934868156925"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C195",
  "IsNew": true,
  "ChemicalName": "Sorbitan monostearate (E491)",
  "Synonyms": [],
  "CASNumber": "1338-41-6",
  "Formula": null,
  "ChemicalGroup": "Chất nhũ hóa / phụ gia thực phẩm",
  "ProductCategoriesText": "Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C196",
  "IsNew": true,
  "ChemicalName": "TEA-dodecylbenzenesulfonate",
  "Synonyms": [],
  "CASNumber": "CAS cần xác minh",
  "Formula": null,
  "ChemicalGroup": "Chất hoạt động bề mặt anion",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân",
  "HazardGroups": "Chưa rà soát phân loại GHS",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "Chưa rà soát",
   "EcoSource": null
  },
  "HPLLevelID": 0,
  "HPLBasis": "Chưa đủ bằng chứng GHS/độc tính để xác định HPL",
  "HPLVerification": "CHƯA PHÂN MỨC",
  "HPLSourceURL": null,
  "CheckNotes": "Cần xác minh GHS/SDS đúng CAS và dạng hóa chất; không tự gán HPL-1.",
  "References": [
   "https://www.pharmacity.vn/dau-goi-clear-9-thao-duoc-co-truyen-lam-sach-gau-350g.html; https://www.sunsilk.com.vn/p/dau-goi-ong-muot-rang-ngoi.html/08934868156925"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C197",
  "IsNew": true,
  "ChemicalName": "Tris[2-(2-(2-methoxyethoxy)ethoxy)ethyl] orthoborate",
  "Synonyms": [],
  "CASNumber": "30989-05-0",
  "Formula": null,
  "ChemicalGroup": "Ester borat / dầu phanh",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H361d"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: Triethylene glycol monomethyl ether borate: nghi ngờ gây độc sinh sản, nhóm 2.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://hcis.safeworkaustralia.gov.au/HazardousChemical/Details?chemicalID=6859",
  "CheckNotes": "Một số thông báo C&L ghi H361fd; dùng H361d từ nguồn phân loại được lựa chọn.",
  "References": [
   "https://msdspds.castrol.com/ussds/amersdsf.nsf/Files/51C1B693F38965B580258C7C00529896/%24File/3180068.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C198",
  "IsNew": true,
  "ChemicalName": "Xanthan gum",
  "Synonyms": [],
  "CASNumber": "11138-66-2",
  "Formula": null,
  "ChemicalGroup": "Polysaccharide / phụ gia thực phẩm",
  "ProductCategoriesText": "Mỹ phẩm và sản phẩm chăm sóc cá nhân; Hóa chất sử dụng trong chế biến và bảo quản thực phẩm",
  "HazardGroups": "Không phân loại GHS theo SDS được lựa chọn",
  "HazardCategories": "Không phân loại",
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "Không phân loại GHS"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": null,
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 1,
  "HPLBasis": "HPL-1: Xanthan gum CAS 11138-66-2: SDS không phân loại nguy hại sức khỏe/sinh thái.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://sdsmanager.com/safety-data-sheet/sigma-aldrich-chemie-gmbh-xanthan-gum-from-xanthomonas-campestris-en/",
  "CheckNotes": "Bụi có thể gây nguy hại cháy nổ; cảnh báo riêng.",
  "References": [
   "https://icheck.vn/san-pham/sua-rua-mat-acnevir-santafa-do-8938535677218; https://richschannel.com/vn/vi/learning/product/bot-banh-bong-lan-richmix-tui-1kg; https://shop.richs.com.vn/kem-lam-banh-rich-s-whip-topping-base-hop-907g; https://www.canova.it/en/mild-and-moderate-acne/287-acnell-crema-restitutiva.html; https://www.pharmacity.vn/dau-goi-clear-9-thao-duoc-co-truyen-lam-sach-gau-350g.html; https://xuanhafood.com.vn/product/kem-lam-banh-richs-niagara-farms/"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 },
 {
  "ChemicalID": "C199",
  "IsNew": true,
  "ChemicalName": "n-Pentane",
  "Synonyms": [],
  "CASNumber": "109-66-0",
  "Formula": null,
  "ChemicalGroup": "Hydrocarbon",
  "ProductCategoriesText": "Nhiên liệu, dung môi và hóa chất kỹ thuật",
  "HazardGroups": "Sức khỏe / Môi trường thủy sinh",
  "HazardCategories": null,
  "Pictograms": [],
  "SignalWord": null,
  "HStatements": [
   "H304",
   "H336",
   "H411"
  ],
  "HealthEffects": null,
  "EnvEffects": null,
  "Ecotox": {
   "LC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "EC50": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "NOEC": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "logKow": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "BCF": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "Biodegradation": "NF – chưa thu thập dữ liệu độc tính sinh thái",
   "AquaticClass": "Aquatic Chronic 2",
   "Assessment": null,
   "DataQuality": "GHS/SDS có nguồn; chưa có đầy đủ LC₅₀/EC₅₀/NOEC",
   "EcoSource": null
  },
  "HPLLevelID": 2,
  "HPLBasis": "HPL-2: n-Pentane: nguy hại hít sặc, STOT SE 3 và Aquatic Chronic 2.",
  "HPLVerification": "CÓ NGUỒN",
  "HPLSourceURL": "https://www.merckmillipore.com/Web-INTL-Site/en_US/-/USD/ShowDocument-File?Country=VN&DocumentId=100882_SDS_VN_VI.PDF&DocumentType=MSD&DocumentUID=3622209&Language=VI&Origin=PDP&ProductSKU=MDA_CHEM-100882",
  "CheckNotes": "H225 dễ cháy là nguy hại vật lý, cần cảnh báo riêng.",
  "References": [
   "https://media.wd40.asia/app/uploads/2024/01/26011053/WD-40-Specialist-Contact-Cleaner-Asia-GHS-SDS-15-11-23.pdf"
  ],
  "StatusNote": "MỚI BỔ SUNG 03/11/2026; chỉ xác minh sự hiện diện trong sản phẩm; chưa thuộc bộ 143 hóa chất đã phân mức HPL."
 }
];
function getChemicalById(id) { return CHEMICALS.find(c => c.ChemicalID === id) || null; }


/* ═══════════════════════════════════════════════════════════════════════
   7. PRODUCT_CHEMICAL — 573 liên kết sản phẩm ⇄ hóa chất
   LabelName = tên thành phần trên nhãn/nguồn; LinkConfidence = độ tin cậy liên kết.
   ═══════════════════════════════════════════════════════════════════════ */
const PRODUCT_CHEMICAL = [
 {
  "ProductID": "P002",
  "ChemicalID": "C053",
  "LabelName": "Ethanol (Cồn thực phẩm)",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P004",
  "ChemicalID": "C062",
  "LabelName": "Acid Clohydric",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P005",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P005",
  "ChemicalID": "C116",
  "LabelName": "Sodium Pyrithione",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P009",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P009",
  "ChemicalID": "C104",
  "LabelName": "Sodium Carbonate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P009",
  "ChemicalID": "C109",
  "LabelName": "Sodium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P010",
  "ChemicalID": "C053",
  "LabelName": "Alcohol Ethoxylate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P010",
  "ChemicalID": "C109",
  "LabelName": "Sodium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P011",
  "ChemicalID": "C028",
  "LabelName": "Cetrimonium Chloride",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P011",
  "ChemicalID": "C055",
  "LabelName": "Etidronic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P011",
  "ChemicalID": "C062",
  "LabelName": "Hydrochloric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P012",
  "ChemicalID": "C109",
  "LabelName": "Sodium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P012",
  "ChemicalID": "C110",
  "LabelName": "Sodium Hypochlorite",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P012",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P012",
  "ChemicalID": "C114",
  "LabelName": "Sodium Metaperriodate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P018",
  "ChemicalID": "C036",
  "LabelName": "Citri Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P018",
  "ChemicalID": "C063",
  "LabelName": "Hydrogen Peroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P018",
  "ChemicalID": "C109",
  "LabelName": "Natri Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P019",
  "ChemicalID": "C063",
  "LabelName": "Hydrogen Peroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P020",
  "ChemicalID": "C044",
  "LabelName": "Didecyldimethylamonium",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P020",
  "ChemicalID": "C045",
  "LabelName": "Diethylene Glycol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P020",
  "ChemicalID": "C053",
  "LabelName": "Ethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P021",
  "ChemicalID": "C044",
  "LabelName": "Didecyldimonium Chloride",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P021",
  "ChemicalID": "C053",
  "LabelName": "Ethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P023",
  "ChemicalID": "C105",
  "LabelName": "NaCl",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P023",
  "ChemicalID": "C109",
  "LabelName": "NaOH",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P023",
  "ChemicalID": "C110",
  "LabelName": "NaOCl",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P024",
  "ChemicalID": "C020",
  "LabelName": "Butane",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P024",
  "ChemicalID": "C053",
  "LabelName": "Cồn/Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P024",
  "ChemicalID": "C093",
  "LabelName": "Propane",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P024",
  "ChemicalID": "C133",
  "LabelName": "Triethanolamine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P026",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C034",
  "LabelName": "CI 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C057",
  "LabelName": "Glucose",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C067",
  "LabelName": "lodopropynyl Butylcarbamate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C068",
  "LabelName": "Lactic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C091",
  "LabelName": "Potassium Sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C096",
  "LabelName": "Propylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C105",
  "LabelName": "Sodium Choride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C118",
  "LabelName": "Sodium Salicylate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C119",
  "LabelName": "Sodium Sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C136",
  "LabelName": "Zinc Pythione",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P027",
  "ChemicalID": "C137",
  "LabelName": "Zinc Sulfate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C029",
  "LabelName": "Cetyl Alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C039",
  "LabelName": "Cocamide MEA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C079",
  "LabelName": "Methylchloroisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C080",
  "LabelName": "Methylispthiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C111",
  "LabelName": "Sodium Lactate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P028",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C018",
  "LabelName": "Biotin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C021",
  "LabelName": "Butylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C023",
  "LabelName": "Calcium Pantothenate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C029",
  "LabelName": "Cetyl Alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C051",
  "LabelName": "DMDM Hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C053",
  "LabelName": "Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C066",
  "LabelName": "Inositol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C067",
  "LabelName": "lodopropynyl Butylcarbamate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C071",
  "LabelName": "Linoleic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C079",
  "LabelName": "Methylchloroisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C080",
  "LabelName": "Methylisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C086",
  "LabelName": "Panthenol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C099",
  "LabelName": "Retinyl Palmitate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C100",
  "LabelName": "Salicylic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C109",
  "LabelName": "Sodium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C111",
  "LabelName": "Sodium Lactate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C121",
  "LabelName": "Sorbitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P029",
  "ChemicalID": "C128",
  "LabelName": "Tocopherol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C037",
  "LabelName": "Climbazole",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C051",
  "LabelName": "DMDM Hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C067",
  "LabelName": "Iodopropynyl Butylcarbamate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C080",
  "LabelName": "Methylisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C100",
  "LabelName": "Salicylic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C109",
  "LabelName": "Sodium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C111",
  "LabelName": "Sodium Lactate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C118",
  "LabelName": "Sodium Salicylate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C136",
  "LabelName": "Zinc Pyrithione",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P030",
  "ChemicalID": "C137",
  "LabelName": "Zinc Sulfate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P031",
  "ChemicalID": "C128",
  "LabelName": "Vitamin E",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl Alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C029",
  "LabelName": "Cetyl Alcool",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C037",
  "LabelName": "Climbazole",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C038",
  "LabelName": "Clotrimazole",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C048",
  "LabelName": "Dipotassium Glycyrrhizate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C054",
  "LabelName": "Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C091",
  "LabelName": "Potassium Sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C096",
  "LabelName": "Propylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C100",
  "LabelName": "Salicylic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C101",
  "LabelName": "Silver",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P032",
  "ChemicalID": "C135",
  "LabelName": "Zinc Oxide",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C029",
  "LabelName": "Cetyl alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C031",
  "LabelName": "Chlorhexidine digluconate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C033",
  "LabelName": "Cl 19140",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C034",
  "LabelName": "Cl 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C086",
  "LabelName": "D-panthenol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P033",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl acetate (Vitamin E)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C022",
  "LabelName": "Canxi Cacbonate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C051",
  "LabelName": "DMDM Hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C057",
  "LabelName": "Glucose",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C072",
  "LabelName": "Lysozyme",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C087",
  "LabelName": "Papain",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C113",
  "LabelName": "Sodium Lauryl Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C115",
  "LabelName": "Sodium Methylparaben",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C117",
  "LabelName": "Sodium Saccharin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C121",
  "LabelName": "Sorbitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P034",
  "ChemicalID": "C124",
  "LabelName": "Tetrasodium Pyrophosphate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P035",
  "ChemicalID": "C033",
  "LabelName": "Cl19140",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P035",
  "ChemicalID": "C053",
  "LabelName": "Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P035",
  "ChemicalID": "C133",
  "LabelName": "Triethanolamine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P036",
  "ChemicalID": "C035",
  "LabelName": "Citral",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P036",
  "ChemicalID": "C041",
  "LabelName": "Coumarin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P036",
  "ChemicalID": "C053",
  "LabelName": "Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C006",
  "LabelName": "Potassium Acesulfame",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C008",
  "LabelName": "Allantoin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C030",
  "LabelName": "Cetylpyridinium Chloride",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C031",
  "LabelName": "Chlorhexidine Digluconate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C032",
  "LabelName": "CI 14720",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C034",
  "LabelName": "CI 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C091",
  "LabelName": "Potassium Sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C121",
  "LabelName": "Sorbitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C126",
  "LabelName": "Thymol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl Acetate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P038",
  "ChemicalID": "C134",
  "LabelName": "Xylitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C016",
  "LabelName": "Benzoic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C053",
  "LabelName": "Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C056",
  "LabelName": "Eucalyptol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C078",
  "LabelName": "Methyl Salicylate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C117",
  "LabelName": "Sodium Saccharin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C121",
  "LabelName": "Sorbitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P039",
  "ChemicalID": "C126",
  "LabelName": "Thymol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C019",
  "LabelName": "Acid boric 3%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C053",
  "LabelName": "Ethanol 96%",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C088",
  "LabelName": "Patent blue",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C107",
  "LabelName": "Natri fluorid",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C117",
  "LabelName": "Natri saccharin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P040",
  "ChemicalID": "C134",
  "LabelName": "Xylitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P049",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P049",
  "ChemicalID": "C054",
  "LabelName": "Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P049",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P049",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P049",
  "ChemicalID": "C094",
  "LabelName": "Propanediol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P050",
  "ChemicalID": "C013",
  "LabelName": "Ascorbic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P050",
  "ChemicalID": "C036",
  "LabelName": "Axit Citric",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P050",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C068",
  "LabelName": "Lactic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C075",
  "LabelName": "Mandelic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C094",
  "LabelName": "Propanediol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P051",
  "ChemicalID": "C106",
  "LabelName": "Sodium Citrate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P052",
  "ChemicalID": "C049",
  "LabelName": "Dipropylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P052",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P052",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C021",
  "LabelName": "Butylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C034",
  "LabelName": "CI 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C053",
  "LabelName": "Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C054",
  "LabelName": "Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C096",
  "LabelName": "Propylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P053",
  "ChemicalID": "C109",
  "LabelName": "Sodium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P054",
  "ChemicalID": "C008",
  "LabelName": "Allantoin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P054",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P054",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P055",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P055",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P055",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate (SLES)",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P056",
  "ChemicalID": "C008",
  "LabelName": "Allantoin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P056",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P056",
  "ChemicalID": "C060",
  "LabelName": "Glycerin / Hyaluronic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P056",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol / Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C001",
  "LabelName": "1,2-Hexanediol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C014",
  "LabelName": "Asiaticoside",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C054",
  "LabelName": "Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C074",
  "LabelName": "Madecassoside",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P057",
  "ChemicalID": "C108",
  "LabelName": "Sodium Hyaluronate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C008",
  "LabelName": "Allantoin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C034",
  "LabelName": "Cl 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C036",
  "LabelName": "Citric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C051",
  "LabelName": "DMDM Hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C061",
  "LabelName": "Hyaluronic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C096",
  "LabelName": "Propylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chlorid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P058",
  "ChemicalID": "C109",
  "LabelName": "Sodium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C012",
  "LabelName": "Arachidic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C026",
  "LabelName": "Carnitine",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C034",
  "LabelName": "Cl 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C053",
  "LabelName": "Alcohol Denat",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C068",
  "LabelName": "Lactic Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C069",
  "LabelName": "Lauric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C081",
  "LabelName": "Myristic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C083",
  "LabelName": "Oleic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C085",
  "LabelName": "Plamitic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C090",
  "LabelName": "Potassium Hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C091",
  "LabelName": "Potassium Sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C113",
  "LabelName": "Sodium Lauryl Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C119",
  "LabelName": "Sodium Sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P059",
  "ChemicalID": "C122",
  "LabelName": "Stearic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P060",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P060",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C017",
  "LabelName": "Benzyl alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C079",
  "LabelName": "Methylchloroisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C080",
  "LabelName": "Methylisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C106",
  "LabelName": "Sodium citrate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P061",
  "ChemicalID": "C113",
  "LabelName": "Sodium lauryl sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C059",
  "LabelName": "Glutathione",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol/Ethylhexylglycerin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P062",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate (SLES)",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P063",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P063",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl Betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P063",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P063",
  "ChemicalID": "C112",
  "LabelName": "Sodium Laureth Sulfate (SLES)",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C002",
  "LabelName": "1-Hydroxyethyl-4,5-diamino pyrazole sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C003",
  "LabelName": "1-Naphthol",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C004",
  "LabelName": "2-Amino-6-chloro-4-nitrophenol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C005",
  "LabelName": "2-Methylresorcinol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C011",
  "LabelName": "Ammonium hydroxide",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C013",
  "LabelName": "Ascorbic acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C073",
  "LabelName": "m-Aminophenol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C109",
  "LabelName": "Sodium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C113",
  "LabelName": "Sodium lauryl sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C120",
  "LabelName": "Sodium sulfite",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P064",
  "ChemicalID": "C130",
  "LabelName": "Toluene-2,5-diamine sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P065",
  "ChemicalID": "C060",
  "LabelName": "Glycerin / Butylene Glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P065",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide 5% (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P066",
  "ChemicalID": "C013",
  "LabelName": "Vitamin C",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P066",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P067",
  "ChemicalID": "C009",
  "LabelName": "Alpha-arbutin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P067",
  "ChemicalID": "C013",
  "LabelName": "Ascorbic Acid (Vitamin C)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P068",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide (Vitamin B3)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P068",
  "ChemicalID": "C131",
  "LabelName": "Tranexamic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P069",
  "ChemicalID": "C011",
  "LabelName": "Ammoium Hydroxide",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P069",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl Alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P069",
  "ChemicalID": "C050",
  "LabelName": "EDTA Disodium Salt",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P069",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P069",
  "ChemicalID": "C120",
  "LabelName": "Sodium Sulfite",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C011",
  "LabelName": "Ammonium hydroxide",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C084",
  "LabelName": "p-Phenylenediamine",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C098",
  "LabelName": "Resorcinol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P070",
  "ChemicalID": "C120",
  "LabelName": "Sodium sulfite",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C001",
  "LabelName": "1,2-Hexanediol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C013",
  "LabelName": "Ascorbic Acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C024",
  "LabelName": "Caprylyl Glycol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C060",
  "LabelName": "Glycerin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C069",
  "LabelName": "Lauric Acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C070",
  "LabelName": "Lauryl Alcohol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C091",
  "LabelName": "Potassium Sobate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C097",
  "LabelName": "Pyridoxine HCl",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C102",
  "LabelName": "Sodium Ascorbyl Phosphate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C103",
  "LabelName": "Sodium Benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C105",
  "LabelName": "Sodium Chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C113",
  "LabelName": "Sodium Lauryl Sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C119",
  "LabelName": "Sodium Sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C126",
  "LabelName": "Thymol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C127",
  "LabelName": "Titanium Dioxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P071",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl Acetate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P072",
  "ChemicalID": "C048",
  "LabelName": "Dipotassium Glycyrrhizate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P072",
  "ChemicalID": "C053",
  "LabelName": "Ethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P072",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P073",
  "ChemicalID": "C042",
  "LabelName": "Cypermethrin 0,2%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P073",
  "ChemicalID": "C065",
  "LabelName": "Imiprothrin 0,05%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P073",
  "ChemicalID": "C132",
  "LabelName": "Transfluthrin 0,06%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P074",
  "ChemicalID": "C042",
  "LabelName": "Cypermethrin 0,1%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P074",
  "ChemicalID": "C065",
  "LabelName": "Imiprothrin 0,06%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P074",
  "ChemicalID": "C132",
  "LabelName": "Transfluthrin 0,15%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P075",
  "ChemicalID": "C015",
  "LabelName": "Atrazine 35%",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P075",
  "ChemicalID": "C077",
  "LabelName": "Mesotrione 10%",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P076",
  "ChemicalID": "C058",
  "LabelName": "Glufosinate Amonium 200g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P077",
  "ChemicalID": "C058",
  "LabelName": "Glufosinate – ammonium 150g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P078",
  "ChemicalID": "C007",
  "LabelName": "Acetaminprid 150g/kg",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P078",
  "ChemicalID": "C047",
  "LabelName": "Dinotefuran 250g/kg",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P079",
  "ChemicalID": "C052",
  "LabelName": "Emamectin benzoate 50g/kg",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P080",
  "ChemicalID": "C042",
  "LabelName": "Cypermthrin 100g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P081",
  "ChemicalID": "C092",
  "LabelName": "Profenofos 450g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P081",
  "ChemicalID": "C125",
  "LabelName": "Thiamethoxam 100g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P082",
  "ChemicalID": "C046",
  "LabelName": "Dimethoate 400g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P083",
  "ChemicalID": "C025",
  "LabelName": "Carbosulfat 200g/l",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P084",
  "ChemicalID": "C064",
  "LabelName": "Imidacloprid 100g/kg",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P085",
  "ChemicalID": "C010",
  "LabelName": "Alphacypermethrin",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P085",
  "ChemicalID": "C043",
  "LabelName": "Deltamethrin",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P096",
  "ChemicalID": "C095",
  "LabelName": "Propanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P099",
  "ChemicalID": "C127",
  "LabelName": "Titanium Dioxide (TiO2)",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P106",
  "ChemicalID": "C138",
  "LabelName": "Chất bảo quản (249)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P106",
  "ChemicalID": "C140",
  "LabelName": "Chất giữ ẩm (451i)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P106",
  "ChemicalID": "C141",
  "LabelName": "Chất giữ ẩm (452i)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P106",
  "ChemicalID": "C142",
  "LabelName": "Chất điều vị (621)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C138",
  "LabelName": "Chất bảo quản (249)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C139",
  "LabelName": "Chất chống oxy hóa (316)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C140",
  "LabelName": "Chất giữ ẩm (451i)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C141",
  "LabelName": "Chất giữ ẩm (452i)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C142",
  "LabelName": "Chất điều vị (621)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P107",
  "ChemicalID": "C143",
  "LabelName": "Phẩm màu tổng hợp (127)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P108",
  "ChemicalID": "C138",
  "LabelName": "Chất bảo quản (249)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P108",
  "ChemicalID": "C139",
  "LabelName": "Chất chống oxy hóa (316)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P108",
  "ChemicalID": "C142",
  "LabelName": "Chất điều vị (621)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "BỘ GỐC – ĐÃ GẮN"
 },
 {
  "ProductID": "P001",
  "ChemicalID": "C101",
  "LabelName": "Silver",
  "HPLScore": 0,
  "HPLText": "CHƯA XÁC MINH LIÊN KẾT",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "TẠM XÁC MINH QUA TÊN SẢN PHẨM"
 },
 {
  "ProductID": "P007",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "SDS ĐÚNG DÒNG, CẦN XÁC NHẬN PHIÊN BẢN"
 },
 {
  "ProductID": "P007",
  "ChemicalID": "C166",
  "LabelName": "Formic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "SDS ĐÚNG DÒNG, CẦN XÁC NHẬN PHIÊN BẢN"
 },
 {
  "ProductID": "P007",
  "ChemicalID": "C159",
  "LabelName": "Deceth-8",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "SDS ĐÚNG DÒNG, CẦN XÁC NHẬN PHIÊN BẢN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C051",
  "LabelName": "DMDM hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C079",
  "LabelName": "Methylchloroisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C080",
  "LabelName": "Methylisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C033",
  "LabelName": "CI 19140",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C191",
  "LabelName": "Sodium alkylbenzene sulfonate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P013",
  "ChemicalID": "C177",
  "LabelName": "Magnesium sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C033",
  "LabelName": "CI 19140",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C088",
  "LabelName": "Patent Blue V",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C191",
  "LabelName": "Sodium alkylbenzene sulfonate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C177",
  "LabelName": "Magnesium sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C172",
  "LabelName": "Hydroxyethyl cellulose",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P014",
  "ChemicalID": "C151",
  "LabelName": "Benzisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C051",
  "LabelName": "DMDM hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C037",
  "LabelName": "Climbazole",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl acetate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C076",
  "LabelName": "Menthol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C106",
  "LabelName": "Sodium citrate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C062",
  "LabelName": "Hydrochloric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C109",
  "LabelName": "Sodium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P041",
  "ChemicalID": "C184",
  "LabelName": "Piroctone olamine",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C091",
  "LabelName": "Potassium sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C067",
  "LabelName": "Iodopropynyl butylcarbamate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C021",
  "LabelName": "Butylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C119",
  "LabelName": "Sodium sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C068",
  "LabelName": "Lactic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C057",
  "LabelName": "Glucose",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C121",
  "LabelName": "Sorbitol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C033",
  "LabelName": "CI 19140",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C034",
  "LabelName": "CI 42090",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C184",
  "LabelName": "Piroctone olamine",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C196",
  "LabelName": "TEA-dodecylbenzenesulfonate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P042",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C127",
  "LabelName": "Titanium dioxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C119",
  "LabelName": "Sodium sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C067",
  "LabelName": "Iodopropynyl butylcarbamate",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C013",
  "LabelName": "Ascorbic acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl acetate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C086",
  "LabelName": "Panthenol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C091",
  "LabelName": "Potassium sorbate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C196",
  "LabelName": "TEA-dodecylbenzenesulfonate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C149",
  "LabelName": "BHT",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P043",
  "ChemicalID": "C194",
  "LabelName": "Sorbic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C081",
  "LabelName": "Myristic acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C069",
  "LabelName": "Lauric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C090",
  "LabelName": "Potassium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C085",
  "LabelName": "Palmitic acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C055",
  "LabelName": "Etidronic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C122",
  "LabelName": "Stearic acid",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C082",
  "LabelName": "Niacinamide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C102",
  "LabelName": "Sodium ascorbyl phosphate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C129",
  "LabelName": "Tocopheryl acetate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C096",
  "LabelName": "Propylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C126",
  "LabelName": "Thymol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C012",
  "LabelName": "Arachidic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C119",
  "LabelName": "Sodium sulfate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C184",
  "LabelName": "Piroctone olamine",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C149",
  "LabelName": "BHT",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P044",
  "ChemicalID": "C173",
  "LabelName": "Hydroxypropyl methylcellulose",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C123",
  "LabelName": "Tetrasodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P045",
  "ChemicalID": "C152",
  "LabelName": "Benzophenone-4",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C039",
  "LabelName": "Cocamide MEA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C068",
  "LabelName": "Lactic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C080",
  "LabelName": "Methylisothiazolinone",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C148",
  "LabelName": "Ammonium lauryl sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C158",
  "LabelName": "Coco-glucoside",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P046",
  "ChemicalID": "C178",
  "LabelName": "Maltodextrin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C112",
  "LabelName": "Sodium laureth sulfate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C040",
  "LabelName": "Cocamidopropyl betaine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C094",
  "LabelName": "Propanediol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C068",
  "LabelName": "Lactic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C036",
  "LabelName": "Citric acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C086",
  "LabelName": "Panthenol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C090",
  "LabelName": "Potassium hydroxide",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C051",
  "LabelName": "DMDM hydantoin",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C048",
  "LabelName": "Dipotassium glycyrrhizate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C050",
  "LabelName": "Disodium EDTA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C157",
  "LabelName": "Cocamide DEA",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C169",
  "LabelName": "Glycolic acid",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C190",
  "LabelName": "Sodium PCA",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P047",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C027",
  "LabelName": "Cetearyl alcohol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C008",
  "LabelName": "Allantoin",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C108",
  "LabelName": "Sodium hyaluronate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C060",
  "LabelName": "Glycerol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C024",
  "LabelName": "Caprylyl glycol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C049",
  "LabelName": "Dipropylene glycol",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C089",
  "LabelName": "Phenoxyethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C103",
  "LabelName": "Sodium benzoate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C154",
  "LabelName": "Betaine",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C162",
  "LabelName": "Dimethicone",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CẦN XÁC MINH",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P048",
  "ChemicalID": "C187",
  "LabelName": "Potassium azeloyl diglycinate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C106",
  "LabelName": "Sodium citrate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C176",
  "LabelName": "Lecithins (E322(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C180",
  "LabelName": "Mono- and diglycerides of fatty acids (E471)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C186",
  "LabelName": "Polysorbate 60 (E435)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C185",
  "LabelName": "Polyglycerol esters of fatty acids (E475)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C193",
  "LabelName": "Sodium stearoyl lactylate (E481(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C173",
  "LabelName": "Hydroxypropyl methylcellulose",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P086",
  "ChemicalID": "C165",
  "LabelName": "Disodium hydrogen phosphate (E339(ii))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C186",
  "LabelName": "Polysorbate 60 (E435)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C193",
  "LabelName": "Sodium stearoyl lactylate (E481(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C175",
  "LabelName": "Lactic acid esters of mono-/diglycerides (E472b)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C180",
  "LabelName": "Mono- and diglycerides of fatty acids (E471)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C179",
  "LabelName": "Methylcellulose (E461)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P087",
  "ChemicalID": "C163",
  "LabelName": "Dipotassium hydrogen phosphate (E340(ii))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C193",
  "LabelName": "Sodium stearoyl lactylate (E481(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C186",
  "LabelName": "Polysorbate 60 (E435)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C180",
  "LabelName": "Mono- and diglycerides of fatty acids (E471)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C185",
  "LabelName": "Polyglycerol esters of fatty acids (E475)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C195",
  "LabelName": "Sorbitan monostearate (E491)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C173",
  "LabelName": "Hydroxypropyl methylcellulose",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C170",
  "LabelName": "Guar gum (E412)",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C167",
  "LabelName": "Gellan gum (E418)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C163",
  "LabelName": "Dipotassium hydrogen phosphate (E340(ii))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P088",
  "ChemicalID": "C153",
  "LabelName": "Beta-carotene (E160a(i))",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C192",
  "LabelName": "Sodium bicarbonate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C164",
  "LabelName": "Disodium diphosphate (E450(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C160",
  "LabelName": "Dicalcium phosphate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C181",
  "LabelName": "Monocalcium phosphate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C168",
  "LabelName": "Glucono delta-lactone",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C185",
  "LabelName": "Polyglycerol esters of fatty acids (E475)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C189",
  "LabelName": "Propylene glycol esters of fatty acids (E477)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C175",
  "LabelName": "Lactic acid esters of mono-/diglycerides (E472b)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C180",
  "LabelName": "Mono- and diglycerides of fatty acids (E471)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P089",
  "ChemicalID": "C198",
  "LabelName": "Xanthan gum",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C105",
  "LabelName": "Sodium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C146",
  "LabelName": "Acetylated distarch adipate (E1422)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C192",
  "LabelName": "Sodium bicarbonate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C164",
  "LabelName": "Disodium diphosphate (E450(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C168",
  "LabelName": "Glucono delta-lactone",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C181",
  "LabelName": "Monocalcium phosphate",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C189",
  "LabelName": "Propylene glycol esters of fatty acids (E477)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C175",
  "LabelName": "Lactic acid esters of mono-/diglycerides (E472b)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P090",
  "ChemicalID": "C193",
  "LabelName": "Sodium stearoyl lactylate (E481(i))",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P091",
  "ChemicalID": "C156",
  "LabelName": "Carrageenan",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P091",
  "ChemicalID": "C174",
  "LabelName": "Konjac glucomannan",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P091",
  "ChemicalID": "C188",
  "LabelName": "Potassium chloride",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P092",
  "ChemicalID": "C147",
  "LabelName": "Agar-agar",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P100",
  "ChemicalID": "C183",
  "LabelName": "Petroleum distillates, hydrotreated light",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P100",
  "ChemicalID": "C155",
  "LabelName": "Carbon dioxide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P101",
  "ChemicalID": "C171",
  "LabelName": "Hexane isomers (mixture)",
  "HPLScore": 0,
  "HPLText": "CHƯA PHÂN MỨC",
  "EvidenceStatus": "CHƯA PHÂN MỨC",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P101",
  "ChemicalID": "C199",
  "LabelName": "n-Pentane",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P101",
  "ChemicalID": "C155",
  "LabelName": "Carbon dioxide",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P102",
  "ChemicalID": "C093",
  "LabelName": "Propane",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P102",
  "ChemicalID": "C020",
  "LabelName": "Butane",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P102",
  "ChemicalID": "C053",
  "LabelName": "Ethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P102",
  "ChemicalID": "C150",
  "LabelName": "Benzalkonium chloride (C12–C16 alkylbenzyl dimethyl ammonium chlorides)",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P103",
  "ChemicalID": "C182",
  "LabelName": "Naphtha (petroleum), hydrotreated light",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CẦN XÁC MINH",
  "LinkConfidence": "SDS ĐÚNG DÒNG, CẦN XÁC NHẬN PHIÊN BẢN"
 },
 {
  "ProductID": "P103",
  "ChemicalID": "C144",
  "LabelName": "2,2'-Methylenebis(6-tert-butyl-4-methylphenol)",
  "HPLScore": 3,
  "HPLText": "HPL-3",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "SDS ĐÚNG DÒNG, CẦN XÁC NHẬN PHIÊN BẢN"
 },
 {
  "ProductID": "P104",
  "ChemicalID": "C022",
  "LabelName": "Calcium carbonate",
  "HPLScore": 1,
  "HPLText": "HPL-1",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "SDS CÙNG DÒNG, CẦN XÁC NHẬN BIẾN THỂ"
 },
 {
  "ProductID": "P105",
  "ChemicalID": "C045",
  "LabelName": "Diethylene glycol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "BỘ GỐC – ĐÃ RÀ SOÁT",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P105",
  "ChemicalID": "C197",
  "LabelName": "Tris[2-(2-(2-methoxyethoxy)ethoxy)ethyl] orthoborate",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P105",
  "ChemicalID": "C145",
  "LabelName": "2-[2-(2-Butoxyethoxy)ethoxy]ethanol",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 },
 {
  "ProductID": "P105",
  "ChemicalID": "C161",
  "LabelName": "Diisopropanolamine",
  "HPLScore": 2,
  "HPLText": "HPL-2",
  "EvidenceStatus": "CÓ NGUỒN",
  "LinkConfidence": "XÁC MINH TỪ THÀNH PHẦN"
 }
];

function getChemicalsByProductId(productId) {
  return PRODUCT_CHEMICAL
    .filter(pc => pc.ProductID === productId)
    .map(pc => ({ ...pc, chemical: getChemicalById(pc.ChemicalID) }))
    .filter(pc => pc.chemical);
}
function getProductsByChemicalId(chemicalId) {
  return PRODUCT_CHEMICAL
    .filter(pc => pc.ChemicalID === chemicalId)
    .map(pc => ({ ...pc, product: getProductById(pc.ProductID) }))
    .filter(pc => pc.product);
}


/* ═══════════════════════════════════════════════════════════════════════
   8. HÀM TIỆN ÍCH DÙNG CHUNG
   ═══════════════════════════════════════════════════════════════════════ */
function normalize(str) {
  return (str || "").toLowerCase().normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "").replace(/đ/gi, "d");
}

function formatFormula(f) {
  if (!f) return "";
  return f.replace(/(\d+)/g, "<sub>$1</sub>");
}

/* Tra cứu hóa chất theo tên (EN/synonyms), công thức hoặc số CAS */
function searchChemicals(query) {
  if (!query || !query.trim()) return [];
  const q = normalize(query.trim());
  return CHEMICALS.filter(c =>
    normalize(c.ChemicalName).includes(q) ||
    (c.Synonyms || []).some(s => normalize(s).includes(q)) ||
    normalize(c.Formula || "").includes(q) ||
    (c.CASNumber || "").includes(query.trim())
  );
}

/* Tra cứu sản phẩm theo tên, nhóm sản phẩm hoặc nhà sản xuất */
function searchProducts(query) {
  if (!query || !query.trim()) return [];
  const q = normalize(query.trim());
  return PRODUCTS.filter(p => {
    const cat = getCategoryById(p.CategoryID);
    return normalize(p.ProductName).includes(q) ||
      normalize(p.Manufacturer || "").includes(q) ||
      (cat && normalize(cat.CategoryName).includes(q));
  });
}

/* Dữ liệu "NF" / "chưa rà soát" → hiển thị mờ, không phải giá trị đo được */
function isDataAvailable(v) {
  if (!v) return false;
  return !/^(NF|N\/A|Chưa\s)/i.test(String(v).trim());
}

/* ── Huy hiệu HPL của HÓA CHẤT (0 = chưa phân mức) ── */
function hplBadgeHTML(levelId, small) {
  const h = getHPLLevel(levelId) || getHPLLevel(0);
  return `<span class="badge" style="${small ? 'font-size:.68rem;' : ''}background:${h.ColorCode}22;color:${h.ColorCode};border-color:${h.ColorCode}55">${h.icon} ${h.HPLCode}</span>`;
}

/* ── HPL của SẢN PHẨM (suy ra từ thành phần) ── */
function productHPLInfo(p) {
  const t = (p.HPLProduct || "CHƯA XÁC ĐỊNH").toUpperCase();
  const m = t.match(/HPL-([123])/);
  const lvl = m ? getHPLLevel(parseInt(m[1])) : getHPLLevel(0);
  const isMin = t.startsWith("TỐI THIỂU");
  let label = m ? (isMin ? `≥ HPL-${m[1]} (tối thiểu)` : `HPL-${m[1]}`) : "Chưa xác định";
  return { level: lvl, label, isMin, known: !!m, icon: m ? lvl.icon : "⚪", color: lvl.ColorCode };
}
function productHPLBadgeHTML(p, small) {
  const i = productHPLInfo(p);
  const style = i.isMin
    ? `background:#fff;color:${i.color};border:1.5px dashed ${i.color}`
    : `background:${i.color}22;color:${i.color};border-color:${i.color}55`;
  return `<span class="badge" style="${small ? 'font-size:.68rem;' : ''}${style}" title="HPL suy ra từ thành phần, không phải phân loại GHS của hỗn hợp">${i.icon} ${i.label}</span>`;
}

/* Giữ tên cũ để tương thích: mức HPL cao nhất đã xác định của sản phẩm (0 nếu chưa xác định) */
function getProductMaxHPL(productId) {
  const p = getProductById(productId);
  return p ? p.HPLScoreMax : 0;
}
