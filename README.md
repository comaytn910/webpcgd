# HỆ THỐNG TRA CỨU & BÁO CÁO PHỔ CẬP GIÁO DỤC - XÓA MÙ CHỮ NĂM 2026
## SỞ GIÁO DỤC VÀ ĐÀO TẠO TỈNH THÁI NGUYÊN
### PHIÊN BẢN WEB APP ĐỘC LẬP (VIEWER MODE) - SẴN SÀNG TRIỂN KHAI CLOUD

---

## 📌 1. Giới thiệu thư mục `webapp/`
Thư mục `webapp/` này được thiết kế độc lập hoàn toàn với phần xử lý dữ liệu nặng:
- **Chỉ chứa dữ liệu kết quả đã tổng hợp:** Không cần thư mục `data/` (92 file nguồn ~250MB) và các processor nặng.
- **Dung lượng siêu nhẹ (~57MB):** Đáp ứng hoàn hảo các giới hạn tải lên của GitHub, Render và Streamlit Cloud.
- **Tốc độ tải tức thì (<0.1 giây):** Nạp thẳng từ các file kết quả Excel, JSON tóm tắt và cơ chế đệm `@st.cache_data`.
- **Đầy đủ 33 biểu nghiệp vụ + Sổ cái Master 33 biểu + Báo cáo đối soát 63 quy tắc.**

---

## 📂 2. Cấu trúc các tệp trong `webapp/`
```text
webapp/
├── .streamlit/
│   └── config.toml          # Cấu hình giao diện và máy chủ Streamlit headless
├── config/
│   └── sheets_catalog.py    # Danh mục chuẩn 33 biểu mẫu nghiệp vụ
├── outputs/                 # Dữ liệu kết quả đã tổng hợp từ đủ 92 xã
│   ├── SO_CAI_TONG_HOP_PCGD_2026_THAI_NGUYEN.xlsx   # Sổ cái 33 biểu cấp tỉnh
│   ├── BAO_CAO_DOI_SOAT_63_QUY_TAC_PCGD_2026.xlsx   # Báo cáo 63 quy tắc logic
│   ├── summary_dashboard.json                       # Tóm tắt chỉ tiêu cốt lõi
│   ├── last_sync_metadata.json                      # Thông tin độ mới dữ liệu
│   └── TONG_HOP_THU_*.xlsx                          # 33 tệp kết quả từng biểu
├── .gitignore               # Loại bỏ các file tạm khi đẩy Git
├── Procfile                 # Lệnh khởi chạy trên Render / Heroku
├── render.yaml              # Cấu hình tự động Blueprint cho Render
├── requirements.txt         # Danh mục thư viện rút gọn (Streamlit, pandas, openpyxl)
├── app.py                   # Ứng dụng web hiển thị chính (Viewer Mode)
├── CHAY_THU_WEBAPP_CUC_BO.bat # Tệp chạy thử trực tiếp trên máy tính
└── README.md                # Tài liệu hướng dẫn chi tiết
```

---

## 🚀 3. Hướng dẫn Đẩy lên GitHub

Người đẹp có thể chọn **1 trong 2 cách** rất đơn giản dưới đây:

### CÁCH A: Kéo thả trực tiếp trên giao diện Web GitHub (Không cần cài đặt Git - Khuyên dùng)
1. Đã có sẵn tệp nén siêu nhẹ chỉ **14.5 MB**: [`webapp_deploy_cloud.zip`](../webapp_deploy_cloud.zip). (Nếu có cập nhật mới, nhấp đúp tệp [`NEN_WEBAPP_DE_DEPLOY.bat`](../NEN_WEBAPP_DE_DEPLOY.bat) để nén lại).
2. Vào [GitHub.com](https://github.com) -> Bấm **New** để tạo một Repository mới (ví dụ đặt tên: `pcgd-thai-nguyen-webapp`, chọn Public).
3. Tại trang vừa tạo, bấm vào dòng chữ: **"uploading an existing file"**.
4. Kéo thả toàn bộ các tệp trong thư mục `webapp` (hoặc các file giải nén từ `webapp_deploy_cloud.zip`) vào khung tải lên của GitHub.
5. Bấm nút màu xanh **"Commit changes"** ở cuối trang. Xong!

---

### CÁCH B: Sử dụng GitHub Desktop hoặc Git dòng lệnh
Nếu máy tính có cài Git:
```powershell
cd "c:\Users\Ad\Desktop\TONG HOP PHO CAP\webapp"
git init
git add .
git commit -m "Khoi tao Web App Tra cuu PCGD 2026 Thai Nguyen"
git branch -M main
git remote add origin https://github.com/<tai-khoan-github-cua-ban>/pcgd-thai-nguyen-webapp.git
git push -u origin main
```

---

## 🌐 4. Hướng dẫn Triển khai (Deploy) Lên Cloud

### LỰA CHỌN 1: Triển khai lên RENDER (Khuyên dùng theo yêu cầu)
Render hỗ trợ Web Service Python cực kỳ mạnh mẽ:

1. Đăng nhập vào [Render.com](https://render.com) (bằng tài khoản GitHub).
2. Bấm nút **New +** ở góc trên bên phải -> Chọn **Web Service**.
3. Chọn kho GitHub `pcgd-thai-nguyen-webapp` vừa đẩy lên.
4. Điền các thông số:
   - **Name:** `pcgd-thai-nguyen`
   - **Region:** Singapore hoặc Oregon (khuyên dùng Singapore để tốc độ tại Việt Nam nhanh nhất).
   - **Branch:** `main`
   - **Runtime:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true`
   - **Instance Type:** Chọn gói **Free** (Miễn phí).
5. Bấm **Create Web Service**.
6. Render sẽ tự động cài đặt thư viện và sau khoảng 2-3 phút, Người đẹp sẽ nhận được đường link web trực tuyến dạng: `https://pcgd-thai-nguyen.onrender.com`.

> **Mẹo nhanh:** Trong thư mục `webapp` đã có sẵn file `render.yaml`. Nếu chọn **New +** -> **Blueprint**, Render sẽ tự động điền toàn bộ cấu hình trên mà không cần gõ phím nào!

---

### LỰA CHỌN 2: Triển khai lên STREAMLIT COMMUNITY CLOUD (Nhanh nhất & Tối ưu nhất)
Streamlit Cloud là nền tảng máy chủ chính hãng của chính Streamlit, **hoàn toàn miễn phí 100%**:
- Máy chủ chạy liên tục không bị "ngủ đông" (sleep) như Render Free.
- Dung lượng RAM cấp sẵn lên tới 1GB (Render Free chỉ có 512MB).
- Tốc độ phản hồi WebSocket thời gian thực mượt mà nhất.

**Cách thực hiện (Chỉ mất 30 giây):**
1. Truy cập [share.streamlit.io](https://share.streamlit.io) và đăng nhập bằng tài khoản GitHub.
2. Bấm nút **New app**.
3. Điền thông tin:
   - **Repository:** `<tai-khoan-github>/pcgd-thai-nguyen-webapp`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Bấm nút **Deploy!**. Hệ thống sẽ cấp ngay một đường link chuyên nghiệp dạng: `https://pcgd-thai-nguyen.streamlit.app`.

---

### ⚠️ Lưu ý kỹ thuật về VERCEL:
- **Vercel** được thiết kế chủ yếu cho các ứng dụng Serverless Function (thích hợp cho Next.js, React hoặc các API trả kết quả ngắn dưới 10 giây). 
- **Streamlit** lại hoạt động dựa trên cơ chế Server kết nối 2 chiều liên tục (WebSocket). Khi deploy Streamlit thuần túy lên Vercel Serverless, người dùng rất dễ gặp hiện tượng mất kết nối WebSocket sau mỗi thao tác click.
- Vì vậy, **Render** (hoặc **Streamlit Community Cloud**) là sự lựa chọn số 1 chuẩn kỹ thuật và ổn định nhất cho ứng dụng này.

---

## 💻 5. Chạy thử nghiệm trực tiếp trên máy tính
Để kiểm tra ứng dụng trước khi đẩy lên mạng, Người đẹp chỉ cần:
1. Vào thư mục `webapp/`.
2. Nhấp đúp vào tệp [`CHAY_THU_WEBAPP_CUC_BO.bat`](CHAY_THU_WEBAPP_CUC_BO.bat).
3. Ứng dụng sẽ tự động mở trên trình duyệt tại địa chỉ: `http://localhost:8502`.

