# H2X Studio — Fee Proposal Generator

## Deploy lên Vercel (5 bước)

### Bước 1: Cài Git & Node.js
- Download Git: https://git-scm.com
- Download Node.js: https://nodejs.org (LTS version)

### Bước 2: Upload code lên GitHub
1. Tạo account GitHub tại github.com (miễn phí)
2. Tạo repository mới, đặt tên `h2x-proposal`
3. Upload toàn bộ thư mục này lên (trừ folder `node_modules`)

### Bước 3: Deploy lên Vercel
1. Vào vercel.com → Sign up bằng GitHub account
2. Click "Add New Project" → chọn repo `h2x-proposal`
3. Vercel tự detect Node.js → Click "Deploy"
4. Chờ ~2 phút → có URL dạng `h2x-proposal.vercel.app`

### Bước 4: Chia sẻ cho team
Gửi URL cho team — ai cũng dùng được, không cần cài gì.

### Bước 5: Cập nhật khi cần
Sửa code → push lên GitHub → Vercel tự deploy lại.

## Cấu trúc file
```
h2x-proposal/
├── server.js          # Express server + AI endpoint
├── generatePptx.js    # PPTX generator (15 slides, H2X brand)
├── fontData.js        # Font H2X embed (auto-generated)
├── package.json       # Dependencies
├── vercel.json        # Vercel config
└── public/
    └── index.html     # Web form UI
```

## Logic màu sắc
- 🔴 ĐỎ = User nhập thông tin (Cover, Project Understanding, Fee, Add-ons, Contact)
- 🔵 XANH = AI đề xuất (Zoning Programme, Design Language)
- ⬜ AUTO = Tự động generate (Scope, Lighting, Team, Process, Terms...)
