# Payroll Hub App

Đây là ứng dụng quản lý, đối soát và báo cáo (Payroll/Audit Hub) được xây dựng bằng React (Vite) và Tailwind CSS.

## Hướng dẫn đưa lên Vercel (Deploy)

Dự án này đã được cấu hình sẵn một file `vercel.json` phục vụ cho React Router (SPA) giúp các đường dẫn (routes) không bị lỗi 404 khi truy cập trực tiếp.

### Cách 1: Deploy qua GitHub (Khuyên dùng)
1. Push toàn bộ source code này lên một repository của bạn trên GitHub, GitLab, hoặc Bitbucket.
2. Đăng nhập vào trang chủ [Vercel](https://vercel.com).
3. Nhấp vào **"Add New..."** > **"Project"** và import repository vừa tạo.
4. Vercel sẽ tự động phát hiện project sử dụng **Vite**. (Nếu không, trong phần *Framework Preset*, hãy chọn **Vite**).
   - *Build Command*: `npm run build` hoặc `vite build`
   - *Output Directory*: `dist`
   - *Install Command*: `npm install`
5. Nhấp **Deploy** và chờ đợi một lát để dự án được public.

### Cách 2: Deploy trực tiếp bằng Vercel CLI
Nếu máy tính của bạn đã cài đặt Node.js và bạn muốn deploy qua Command Line:

1. Mở terminal và cài đặt Vercel CLI:
   ```bash
   npm i -g vercel
   ```
2. Đăng nhập vào Vercel từ terminal:
   ```bash
   vercel login
   ```
3. Đứng ở thư mục gốc của project này và chạy lệnh:
   ```bash
   vercel
   ```
4. Xác nhận các tùy chọn hiển thị trên terminal (nhấn Enter cho các giá trị mặc định). Vercel sẽ tải dự án hiển thị cho bạn đường link dự án. 
   - Để deploy bản chính thức (production), chạy tiếp lệnh:
   ```bash
   vercel --prod
   ```

## Development
- `npm install` để cài đặt thư viện.
- `npm run dev` để chạy môi trường code.
- `npm run build` để đóng gói sinh ra file chạy production.

## Đồng bộ dữ liệu giữa các máy

Dùng cùng tài khoản Payroll để mở số liệu và các tháng đã lưu trên máy khác. Xem [cách lưu, đồng bộ và xử lý lỗi](docs/cloud-workspace.md), bao gồm kiểm thử và giới hạn xác minh.

## Lưu ý Quan trọng khi Deploy lên Vercel và Đồng bộ Supabase

Khi deploy ứng dụng lên Vercel, nếu bạn gặp lỗi không thể lưu dữ liệu hay đồng bộ (sync) lên Supabase, nguyên nhân phổ biến nhất là **thiếu hoặc chưa cấu hình đúng các Biến Môi trường (Environment Variables)** trên bảng điều khiển (Dashboard) của Vercel.

### Các bước khắc phục lỗi kết nối Supabase trên Vercel:

1. **Đăng nhập vào Dashboard của Vercel**:
   - Truy cập [Vercel](https://vercel.com) và chọn dự án (project) của bạn.

2. **Truy cập phần cài đặt biến môi trường**:
   - Vào mục **Settings** -> **Environment Variables**.

3. **Thêm đầy đủ 2 biến môi trường Client-side của Supabase**:
   - **Tên biến**: `VITE_SUPABASE_URL`
     - **Giá trị**: URL dự án Supabase của bạn (ví dụ: `https://xxxx.supabase.co`).
   - **Tên biến**: `VITE_SUPABASE_ANON_KEY`
     - **Giá trị**: Khóa anon public key của dự án Supabase.
   - *Lưu ý*: Phải giữ nguyên tiền tố `VITE_` ở đầu tên biến thì ứng dụng React mới có thể đọc được ở phía client-side.

4. **Thêm các biến môi trường Server-side (nếu dùng tính năng đồng bộ Google Drive/Sheets)**:
   - Thêm `GOOGLE_CLIENT_EMAIL` và `GOOGLE_PRIVATE_KEY` vào cùng mục Settings của Vercel nếu bạn sử dụng các tính năng đồng bộ dữ liệu tự động từ tài khoản Google Drive dịch vụ.

5. **Redeploy lại dự án**:
   - Sau khi lưu các biến môi trường, hãy vào mục **Deployments** trên Vercel, chọn lượt deploy gần nhất và bấm **Redeploy** (hoặc tạo một commit mới và push lên GitHub) để Vercel build lại ứng dụng với các cấu hình biến môi trường mới nhất.

### Kiểm tra trạng thái trên giao diện ứng dụng:
- Sau khi redeploy thành công, mở trang web của bạn trên Vercel.
- Kiểm tra góc giao diện hoặc thông báo xem hệ thống có còn báo lỗi `"Supabase configuration is missing or invalid"` hay không. Nếu các biến được nhận đúng, bạn sẽ đồng bộ được dữ liệu lên Supabase bình thường!

