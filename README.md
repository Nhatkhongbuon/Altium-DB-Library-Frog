# Altium DB Library Helper

**Phiên bản 1.7.6** — Nhatnpm

Tool chạy trên PC: nhập **part number** → chọn (hoặc để tool tự nhận diện) **loại linh kiện** → tra cứu Digi-Key / Mouser → điền đúng các trường của sheet tương ứng → ghi vào file Excel dùng làm database của Altium.

Từ **1.7.3** tool đọc và ghi được **hai kiểu file database** — tự nhận diện khi mở file, và chuyển đổi qua lại được:

- **`Master_Library.xlsx` (From Nhatnpm)** — **kiểu mặc định**, khoá `Design Item ID`
- **`Database_Library.xlsx` (From Hoangnh277)** — **kiểu tuỳ chọn**, khoá `Mfr. #`, dữ liệu nằm trong Excel Table

Xem mục **3k**.

Từ **1.7.4**, mọi **bảng** trong tool đều **bôi đen và copy được** — Ctrl+C, Ctrl+A, menu chuột phải, và lớp phủ để bôi đen từng đoạn chữ ngay trong ô. Xem mục **3l**.

Từ **1.7.5**: kéo thả đổi kích thước cửa sổ **mượt hẳn**, nút trên **thanh taskbar đeo đúng icon ếch / nhện** thay vì icon Python, **đóng gói được thành file `.exe`** chạy độc lập không cần cài Python, và **bản `.exe` tự cập nhật được** từ GitHub Release. Xem mục **1b**, **3m**, **3n**.

Từ **1.7.6**: ô **Repo GitHub** trong Cấu hình bị **ẩn sau mật khẩu** — xem mục **3i**.

---

## 1. Cài đặt

1. Cài Python 3.9+ từ https://www.python.org/downloads/ — nhớ tick **"Add Python to PATH"**.
2. Giải nén thư mục này ra đâu cũng được (ví dụ `D:\Altium\Tools\AltiumDbTool`).
3. Double-click **`run.bat`**. Lần đầu nó tự cài `requests`, `openpyxl`, `xlwings` rồi mở giao diện.

### Ba cách mở tool

| Cách | Cửa sổ đen | Dùng khi nào |
|---|---|---|
| **`run_silent.vbs`** | không hiện gì cả | dùng hằng ngày — khuyến nghị |
| **`AltiumDbTool.pyw`** | không hiện | double-click quen tay kiểu file Python |
| **`run.bat`** | nháy một cái rồi tắt | lần đầu (cần cài thư viện) hoặc khi muốn xem lỗi cài đặt |

`run.bat` giờ khởi động tool bằng `pythonw.exe` rồi tự thoát, nên cửa sổ Command Prompt không ở lại nữa. Muốn không thấy nó nháy chút nào thì dùng `run_silent.vbs`.

Tiện nhất: nháy phải `run_silent.vbs` → **Send to → Desktop (create shortcut)**, đổi tên shortcut cho gọn rồi ghim vào Taskbar.

Cài tay nếu muốn:

```
pip install -r requirements.txt
python altium_db_tool.py
```

## 1b. Đóng gói thành file `.exe`  *(1.7.5)*

Muốn chạy tool mà không cần cài Python, hoặc muốn ghim một icon đẹp lên Taskbar:
**double-click `build_exe.bat`**. Script tự cài `PyInstaller` nếu thiếu, dọn thư
mục build cũ, build, rồi đóng gói luôn bản phát hành.

| File | Là gì |
|---|---|
| `dist\AltiumDbTool.exe` | file chạy, icon con ếch |
| `dist\version.json` | số phiên bản của file exe cạnh nó |
| `..\History\AltiumDbTool_v<ver>_exe.zip` | gói để đính kèm lên GitHub Release |

**Chỉ một file exe, icon con ếch.** Icon Windows vẽ cho một `.exe` trên Explorer và
trên nút Taskbar là icon **nhúng sẵn bên trong chính file đó** — Windows đọc nó từ
lúc chưa mở chương trình, nên nó không thể đọc `config.json`. Một exe chỉ nhúng
được một icon.

Lựa chọn **Ếch / Nhện bên trong tool vẫn giữ nguyên** (Functions → Giao diện →
Logo): nó đổi icon **cửa sổ**, icon **Taskbar lúc đang chạy**, và ảnh ở tab Liên
hệ. Chỉ riêng icon của file `.exe` trên Explorer là cố định con ếch.

Chép `AltiumDbTool.exe` sang thư mục làm việc rồi chạy. `config.json`,
`rules.json`, `api_keys.enc`, `lookup_cache.db`, `state_log.db`, `version.json`
được tạo **cạnh file `.exe`** — đổi chỗ thì nhớ chép chúng theo.

### Phát hành lên GitHub

Sau khi build, `build_exe.bat` gọi `make_release.py` để sinh `version.json` và nén
`History\AltiumDbTool_v<ver>_exe.zip` đúng cấu trúc mà updater đợi:

```
AltiumDbTool/AltiumDbTool.exe
AltiumDbTool/version.json
```

Đem file `.zip` đó lên **GitHub → Releases → Draft a new release**, tag đặt kiểu
`v1.7.5`, và thả file vào khung **"Attach binaries by dropping them here"** ở mục
**Assets** — *không* phải ô soạn ghi chú. Cách tool tự cập nhật bản exe: xem mục
**3n**.

Muốn tự chạy lệnh thì:

```
pip install pyinstaller
pyinstaller AltiumDbTool.spec --noconfirm --clean
python make_release.py
```

Ghi chú:

- Các file ảnh (`frog.ico`, `frog.png`, `frog_icon.png`, `icon.ico`, `icon.png`)
  được nhúng vào gói; hàm `resource()` trong `altium_db_tool.py` tự tìm chúng
  trong thư mục giải nén tạm của PyInstaller (`sys._MEIPASS`).
- `config.json` / `rules.json` / `api_keys.enc` **không** nhúng vào exe — chúng là
  file của bạn, phải nằm cạnh exe để sửa được và còn giữ lại sau khi cập nhật.
  Thiếu thì tool tự tạo lại từ mặc định.
- UPX bị **tắt** trong `AltiumDbTool.spec`: file nén UPX hay bị phần mềm diệt virus
  báo nhầm là mã độc.
- `version.json` là thứ tool đọc để biết **file exe trên đĩa đang là bản nào**.
  Không đọc được số đó từ chính file exe: mã nguồn bên trong đã thành `.pyc` nên
  không còn dòng chữ `APP_VERSION = "1.7.5"` nào để tìm.
- `version_info.txt` là thứ khác: thông tin hiện trong *Properties* của file exe
  (tên, phiên bản, tác giả). Lên phiên bản mới nhớ sửa số ở **cả bốn chỗ**.
- `APP_DIR` khi chạy exe lấy từ `sys.executable` (không phải `sys.argv[0]`), nên
  chạy từ shortcut có đặt "Start in" khác chỗ vẫn tìm đúng `config.json`.

## 2. Lấy API key

### Digi-Key (khuyến nghị — dữ liệu khớp nhất với file của bạn)

1. Đăng ký tài khoản tại https://developer.digikey.com
2. **My Organizations → Create Organization** → **Create Production App**
3. Chọn API **Product Information V4**, ghi lại **Client ID** và **Client Secret**
4. Dán vào tool: nút **Cấu hình API**

### Mouser (đăng ký nhanh hơn, dùng làm nguồn dự phòng)

1. https://www.mouser.vn/api-search/ → đăng ký **Search API**
2. Nhận API key qua email → dán vào ô Mouser trong **Cấu hình API**

### Nạp key từ file TXT

Thay vì gõ tay, bấm **Nạp key từ file TXT...** trong hộp thoại Cấu hình API và chọn file dạng:

```
Digikey Client ID: <client id của bạn>
Digikey Client Secret: <client secret của bạn>

Mouser API: <api key của bạn>
```

Không phân biệt hoa thường, chấp nhận cả `Digi-Key Client ID`, `Mouser API Key`, `Mouser Key`. Dòng trống và dòng bắt đầu bằng `#` bị bỏ qua.

Ngoài ra, nếu đặt file tên **`API.txt`** ngay cạnh `altium_db_tool.py`, lần chạy đầu tool tự nạp khi trong `config.json` chưa có key nào.

### Từ 1.6: key được **mã hoá**, không còn nằm trong `config.json`

Bản 1.5.x lưu Client ID / Secret / Mouser key thẳng vào `config.json` dưới dạng chữ thường — mà `config.json` lại nằm trong repo Git, ai clone về là cầm được key.

Bản 1.6 cất key vào một file riêng đã mã hoá, mặc định `api_keys.enc`, theo một trong hai cách chọn trong **Cấu hình → Bảo mật API key**:

| Chế độ | Cần mật khẩu? | Mang sang máy khác | Dùng khi nào |
|---|---|---|---|
| **DPAPI** (khuyến nghị) | không | không (khoá gắn với tài khoản Windows) | máy cá nhân, dùng hằng ngày |
| **Mật khẩu** | có, mỗi phiên một lần | được | dùng chung nhiều máy, hoặc để dành import lại |
| Không mã hoá | không | được | chỉ khi bạn chắc thư mục không lên Git |

Cách dùng thường gặp:

1. Nhập key ở thẻ **API** → sang thẻ **Bảo mật API key** → chọn chế độ → bấm **Mã hoá & lưu ngay**.
2. Muốn dùng lại ở máy khác: bấm **Xuất file key...** (chế độ mật khẩu) → chép file `.enc` sang máy kia → **Nạp file key...** → nhập mật khẩu. Không phải gõ lại key bao giờ nữa.
3. Lần chạy đầu, nếu `config.json` còn key dạng chữ thường, tool hỏi ngay *"Mã hoá bây giờ?"* — đồng ý là nó mã hoá, xoá key khỏi `config.json` và thêm `config.json` / `api_keys.enc` / `API.txt` vào `.gitignore` giúp luôn.

Kỹ thuật: DPAPI gọi thẳng `CryptProtectData` của Windows; chế độ mật khẩu dùng PBKDF2-HMAC-SHA256 200.000 vòng để sinh khoá, mã hoá kiểu CTR dựa trên SHA-256 và xác thực bằng HMAC-SHA256 — chỉ dùng thư viện chuẩn của Python, không cài thêm gói nào. Nhập sai mật khẩu là báo sai chứ không trả ra dữ liệu rác.

> Nếu repo từng public với key cũ trong đó thì vẫn nên **thu hồi và xin key mới** — key cũ đã nằm trong lịch sử Git rồi, mã hoá bây giờ không xoá được quá khứ.

Chọn nguồn ở ô **Nguồn**: `Digi-Key`, `Mouser`, hoặc `Digi-Key + Mouser` (tra Digi-Key trước, thiếu trường nào thì lấy bù từ Mouser).

## 3. Dùng

1. **Chọn...** → trỏ tới `Master_Library.xlsx` (tool nhớ đường dẫn cho lần sau). Chưa có file thì bấm **Tạo mới...** — xem mục 3j.
2. Gõ part number → **Enter** hoặc bấm **Tra cứu**.
3. Bảng giữa hiện tất cả các trường của sheet đó:
   - dòng **đen** = đã có dữ liệu
   - dòng **đỏ** = còn trống, nên tự điền
   - dòng **xanh** = cột do **công thức Excel** tự tính (Description, Comment, Footprint Ref, Value Sort... tuỳ sheet)
   - Nháy đúp vào ô **Giá trị** để sửa tay trước khi ghi.
4. **Ghi vào Excel**. Nếu part number đã có, tool hỏi có cập nhật đè dòng cũ không.

## 3a. Chọn nhà sản xuất

Cùng một part number có thể do nhiều hãng làm, thông số lệch nhau đôi chút. Khi tra cứu ra **nhiều hơn một kết quả**, tool mở cửa sổ chọn:

| ✓ | Part number | Nhà sản xuất | Package | In-Stock | Mô tả | Nguồn |
|---|---|---|---|---|---|---|
| ☑ | RMCF0805FT4K70 | Stackpole Electronics | 0805 | 183.492 | RES 4.7K OHM 1% 1/8W | Digi-Key |
| ☐ | RC0805FR-074K7L | YAGEO | 0805 | 0 | RES 4.7K OHM 1% 1/8W | Digi-Key |
| ☐ | ERJ-6ENF4701V | Panasonic | 0805 | 24.500 | RES 4.7K OHM 1% 1/8W | Mouser |

- Dòng **xanh lá** = trùng khớp chính xác part number bạn gõ, luôn xếp lên đầu
- Dòng **đỏ** = In-Stock bằng 0 (hết hàng, hoặc nhà cung cấp không trả về số liệu)
- **Nháy đúp** hoặc **Xem mục đang chọn** → đưa một hãng vào bảng chính để xem đủ trường và sửa tay
- Tích nhiều dòng rồi **Ghi các mục đã chọn** → chuyển sang bảng xem trước hàng loạt, mỗi hãng thành một dòng riêng trong Excel
- **Chọn tất cả** nếu muốn nạp hết mọi hãng vào thư viện

Chỉ có một kết quả thì tool vào thẳng bảng chính như cũ. Số kết quả tối đa chỉnh trong `config.json` ở khoá `max_results` (mặc định 10).

## 3c. In-Stock

Mỗi lần tra cứu, tool lấy luôn số lượng còn trong kho:

- **Dưới ô tra cứu**: còn hàng hiện chữ xanh `In-Stock: 183.492  (Digi-Key)   •   đặt tối thiểu 1`, hết hàng hiện chữ đỏ `In-Stock: 0  —  hết hàng`
- **Bảng chọn nhà sản xuất** có cột *In-Stock*, dòng bằng 0 tô đỏ để nhìn phát biết ngay
- Số liệu lấy từ `QuantityAvailable` (Digi-Key) và `AvailabilityInStock` (Mouser). Ở chế độ `Digi-Key + Mouser`, hãng nào có số thì lấy số đó.

**Ghi In-Stock vào Excel:** chỉ khi sheet của bạn có sẵn một cột tên `In-Stock`, `Stock`, `Quantity Available`, `Tồn kho` hoặc `Available`. Không có cột thì tool chỉ hiển thị, không tự thêm cột. Hết hàng thì ghi số **0** chứ không để trống, để lọc trong Excel cho dễ.

Lưu ý In-Stock thay đổi từng giờ nên con số ghi vào chỉ là ảnh chụp lúc thêm linh kiện — dùng tham khảo khi chọn hãng, đừng coi là dữ liệu mua hàng.

## 3b. Thêm hàng loạt

Bấm **Thêm hàng loạt** → dán danh sách, mỗi dòng một part number → **Tra cứu**. Tool tra cứu tất cả nhưng **chưa ghi gì**, rồi mở bảng xem trước:

| Cột | Ý nghĩa |
|---|---|
| ✓ | Nháy vào để chọn/bỏ chọn mã đó |
| Part number / Sheet / Nhà sản xuất / Mô tả | Kết quả tra cứu |
| Trạng thái | `Sẵn sàng`, `Sẵn sàng (n trường trống)`, `Đã có - sẽ bỏ qua` (cam), `Chưa rõ loại` / `Lỗi: ...` (đỏ), `Đã ghi vào ... dòng n` (xanh) |

- **Nháy đúp** vào một dòng → mã đó hiện ra ở cửa sổ chính để xem đủ mọi trường, sửa tay, hoặc đổi sheet. Sửa xong quay lại bảng là thấy cập nhật ngay. Bấm **Ghi vào Excel** ở cửa sổ chính thì chỉ ghi riêng mã đang xem.
- Mã **chưa rõ loại**: nháy đúp rồi chọn sheet ở ô **Loại** — bảng tự điền và mã đó được chọn lại.
- Tick **Ghi đè mã đã có trong thư viện** nếu muốn cập nhật các mã trùng, không tick thì chúng bị bỏ qua.
- **Ghi các mã đã chọn** → ghi lần lượt, trạng thái từng dòng cập nhật theo thời gian thực.

Mã trùng nhau trong danh sách nhập vào được tự động gộp lại.

## 3g. Cảnh báo trùng ngay lúc thêm  *(1.6)*

Tra cứu xong, trước khi bạn bấm Ghi, tool tự soi thư viện xem **đã có mã nào cùng bộ thông số chưa** (cùng giá trị, package, dung sai, công suất...). Có thì hiện dòng cam:

```
⚠ Đã có 2 mã cùng thông số: RL0402FR-070R5L (YAGEO, dòng 3), PT0402FR-070R5L (YAGEO, dòng 4) — cùng Package=0402 | Comment=0.5R | Tolerance=1%
```

Bấm **Ghi vào Excel** thì tool hỏi lại lần nữa, liệt kê đủ các mã trùng để bạn quyết định: dùng mã có sẵn hay vẫn thêm mã mới.

Đây là chỗ chặn thư viện phình hiệu quả nhất — tab *Trùng thông số* chỉ là hậu kiểm, lúc mã đã nằm trong file rồi. Tắt được ở **Cấu hình → Giao diện & thói quen**.

## 3d. Bộ công cụ thư viện

Toàn bộ giao diện chia thành **các tab kiểu trình duyệt nằm trên cùng cửa sổ** — không còn thanh menu. Bấm tab là đổi vùng làm việc ngay trong cửa sổ chính, không mở thêm cửa sổ nào:

`Thêm linh kiện` · `Tra cứu thư viện` · `Quét State` · `Sức khoẻ thư viện` · `Symbol / Footprint` · `Trùng thông số` · `Đề xuất thay thế` · `Đối chiếu BOM` · `Giá BOM` · `Where-used` · `Thống kê` · `Liên hệ`

Dòng File Excel dùng chung cho mọi tab. Từ **1.6.4** thanh này chỉ còn 4 nút hay dùng nhất — **Chọn... · Tạo mới... · Nạp lại · Mở Excel** — cộng nút **⚙ Functions** và nút đổi sáng/tối.

### Cửa sổ Functions  *(1.6.4)*

Mọi chức năng phụ gom vào một cửa sổ: cột trái là danh sách, bấm vào mục nào thì nội dung hiện ở cột phải, kiểu File Explorer của Windows.

| Nhóm | Mục | Nội dung |
|---|---|---|
| Thư viện | **Backup & khôi phục** | Danh sách bản backup, khôi phục, so sánh với file hiện tại, dọn bớt |
| Thư viện | **History của Altium** | Dọn `History\*.Zip` trong `Symbols\` và `Footprints\` |
| Thư viện | **Chuyển đổi định dạng DB** *(1.7.3)* | Đổi file đang mở qua lại giữa kiểu mặc định `Master_Library` *(From Nhatnpm)* và kiểu tuỳ chọn `Database_Library` *(From Hoangnh277)* — xem trước rồi ghi ra file mới, xem mục **3k** |
| Cài đặt | **Cấu hình** | 4 thẻ: API · Thư mục & ghi file · Bảo mật API key · Cache / Giao diện / Cập nhật. Nút **Lưu** nằm ở thanh dưới cửa sổ *(1.7.2)* |
| Cài đặt | **Ghim tab** | Chọn tab nào hiện trên thanh làm việc |
| Cài đặt | **Giao diện** | Đổi nhanh Sáng / Tối / Theo Windows, và chọn logo Nhện / Ếch *(1.6.9)* |

Chỉ mở **một** cửa sổ Functions: đang mở mà bấm lại thì nó nhảy tới đúng mục và nổi lên trước. Cửa sổ này **không chiếm quyền điều khiển**, nên mở bảng so sánh backup từ trong đó vẫn bấm được bình thường.

Từ **1.6.5** cả cửa sổ chỉ còn **một** nút **Đóng** — nút ở góc dưới bên phải cửa sổ. Trước đó mỗi panel mang theo nút Đóng riêng của nó, nên khi nhúng vào cửa sổ Functions màn hình hiện ra hai nút Đóng chồng nhau. Phím **Esc** vẫn đóng cửa sổ như cũ.

**Thanh cuộn *(1.6.6, sửa lại 1.6.7)*.** Vùng nội dung bên phải nằm trong một khung cuộn được: trang **Cấu hình** dài hơn màn hình, hay cửa sổ bị kéo nhỏ lại, thì cuộn xuống chứ không còn bị cắt mất phần dưới (trước đây ô *Repo GitHub* và *GitHub token* hay bị khuất). Lăn **bánh xe chuột** ở bất kỳ đâu trong trang là cuộn — trừ khi con trỏ đang nằm trên bảng, ô văn bản hay hộp chọn, những thứ tự cuộn lấy. Nhờ có thanh cuộn nên cửa sổ thu nhỏ được tới **660×420** (trước là 900×520).

Từ **1.6.7** cả hai thanh cuộn **luôn nằm đó**, giống mọi cửa sổ khác của Windows: hết nội dung để cuộn thì con trượt chiếm hết đường ray, nhìn là biết đang ở cuối trang. Bản 1.6.6 cho thanh ngang tự ẩn tự hiện, bố cục nhảy một cái mỗi lần nó xuất hiện.

Cũng từ 1.6.7, khung cuộn **tự đo lại kích thước mỗi 1/4 giây**. Lý do: khung bên trong bị ghim cứng kích thước nên khi nội dung phình ra — chữ xuống dòng thêm, bảng thêm vài hàng — kích thước *thật* của nó không đổi, Tk không báo sự kiện nào, và vùng cuộn giữ nguyên số cũ: con trượt chiếm hết đường ray trong khi phần dưới trang bị cắt, phải kéo cửa sổ một cái thanh cuộn mới chịu đúng. Nay tự sửa lấy.

**Chữ tự xuống dòng *(1.6.6)*.** Các dòng mô tả dài trước đây đặt bề rộng xuống dòng bằng một con số pixel cố định, thu nhỏ cửa sổ là mất chữ. Nay chúng tự tính lại bề rộng mỗi khi khung đổi kích thước, nên chữ luôn nằm gọn trong khung. Áp dụng cho cả cửa sổ **Tạo file thư viện mới** — cửa sổ này giờ cũng có thanh cuộn, hàng nút *Tạo file... / Huỷ* ghim ở đáy, và thu nhỏ được tới 640×460.

### Tra cứu thư viện

Xem và tìm trong toàn bộ database ngay trong app, không phải mở Excel.

- Gõ từ khoá — tìm trên **mọi cột** (nhiều từ = phải khớp tất cả)
- Lọc theo sheet, hoặc tick **Chỉ hiện mã có vấn đề**
- **Bộ lọc tham số (1.6)** — hàng lọc thứ hai:
  - **State** và **Package**: combo tự nạp đúng các giá trị đang có trong sheet đang chọn
  - **Trị số từ … đến …**: gõ `1k` → `10k` (điện trở) hoặc `100n` → `10u` (tụ). Tool đọc cột `Value Sort`, không có thì tự quy đổi từ `Comment` theo đơn vị khai trong `rules.json`
  - **Cột … chứa …**: chọn bất kỳ cột nào của sheet rồi lọc theo chuỗi
  - **Xoá lọc** trả mọi thứ về mặc định
- **Đỏ** = NRND / Obsolete / Last Time Buy — **Cam** = In-Stock = 0 / Not Found — **Xanh** = Normal
- Chọn một dòng, bảng dưới hiện **toàn bộ thông số** của dòng đó
- Nút **Đề xuất thay thế** và **Tải datasheet** cho mã đang chọn

### Quét State  *(1.5 gọi là "Kiểm tra thư viện")*

Quét toàn bộ thư viện, tra cứu từng mã rồi cập nhật cột **State**:

| State | Khi nào |
|---|---|
| `Normal` | còn hàng, vòng đời bình thường |
| `NRND` | Not Recommended for New Designs |
| `Last Time Buy` | đợt mua cuối |
| `Obsolete` | ngừng sản xuất / discontinued |
| `In-Stock = 0` | còn sống nhưng hết hàng |
| `Not Found` | nhà cung cấp không còn mã này |

Vòng đời được ưu tiên hơn tồn kho — mã Obsolete mà vẫn còn hàng thì vẫn ghi `Obsolete`.

Có thanh tiến độ, nút **Dừng**, và ô **nghỉ giữa các lần gọi** để không đụng giới hạn API. Kết quả hiện ra bảng trước, bạn xem xong mới bấm **Ghi cột State vào Excel** — không tự ghi bừa. Xuất CSV được.

Linh kiện thêm mới cũng tự có State ngay lúc ghi.

**Cache (1.6):** mã nào vừa tra trong vòng 12 giờ thì lấy thẳng từ cache, không gọi API và không phải nghỉ giữa các lần gọi — quét lần thứ hai trong ngày gần như tức thì. Tick **Bỏ qua cache (ép gọi API)** khi cần số tồn kho mới tinh.

**Nhật ký thay đổi State (1.6):** mỗi lần quét xong, tool tự lưu một ảnh chụp và so ngay với lần quét trước:

```
So với lần quét 20/08/2026 09:14: 4 mã đổi trạng thái (2 mã xấu đi: TPS62933FDRLR Normal→NRND, RC0402FR-071KL Normal→In-Stock = 0)  •  1 mã mới
```

Bấm **Nhật ký State...** để mở cửa sổ đầy đủ: danh sách các lần quét, chọn hai lần bất kỳ (giữ Ctrl) để so, dòng **đỏ** = xấu đi, **xanh** = tốt lên, xuất CSV được. Dữ liệu nằm trong `state_log.db`, mặc định giữ 50 lần quét gần nhất.

### Sức khoẻ thư viện  *(1.6)*

Tab này **không cần mạng** — nó soi lỗi ngay trong dữ liệu Excel và các file thư viện:

| Phát hiện | Mức độ |
|---|---|
| `Design Item ID` trùng trong cùng sheet / có ở hai sheet / bỏ trống / thừa khoảng trắng | Lỗi · Cảnh báo |
| Sheet thiếu cột bắt buộc (`Design Item ID`, `Library Path/Ref`, `Footprint Path/Ref`) | Lỗi |
| Sheet thiếu cột `State` | Cảnh báo |
| Tiêu đề cột chứa **ký tự vô hình** — non-breaking space, khoảng trắng thừa, tab | Cảnh báo |
| Hai cột cùng tên | Lỗi |
| `Footprint Ref` khác quy luật mà chính sheet đó đang dùng cho package ấy | Cảnh báo |
| `Value Sort` không khớp trị số (tính lại theo `rules.json`) | Cảnh báo |
| Link datasheet thiếu `https:` (Digi-Key hay trả `//mm.digikey.com/...`) hoặc không hợp lệ | Gợi ý · Cảnh báo |
| Thiếu Manufacturer / Description / Comment / link datasheet | Gợi ý |
| **Symbol / footprint mồ côi**: có trong `.SchLib`/`.PcbLib` nhưng không dòng Excel nào trỏ tới | Gợi ý |
| **Footprint chưa gắn 3D model** | Gợi ý |

Tick từng mức độ để lọc, nháy đúp một dòng để nhảy sang tab *Tra cứu thư viện* với đúng mã đó, xuất CSV được.

Ngay đầu tab còn có khung **File .DbLib** — xem mục 3h.

### Tự sửa đường dẫn `.DbLib`  *(1.6)*

Khung trên cùng của tab *Sức khoẻ thư viện*. Tool tự tìm mọi file `.DbLib` dưới gốc thư viện, và mỗi lần bấm **Kiểm tra lại** (hoặc đổi file trong ô File) nó đọc lại file đó rồi đối chiếu với máy hiện tại:

- `Data Source=` đang trỏ tới đâu, file đó có tồn tại không, có đúng file Excel đang mở trong tool không
- có bao nhiêu bảng đang **Enabled**, sheet nào trong Excel **chưa** được bật trong `.DbLib`, bảng nào thừa
- đã cài **Microsoft.ACE.OLEDB** chưa và là bản **32 hay 64 bit** (đọc registry) — Altium 64-bit mà chỉ có ACE 32-bit là không kết nối được

Nút **Sửa đường dẫn theo máy này** chỉ sáng khi `Data Source` đang sai — đúng rồi thì nó xám, không có gì phải làm. Sai đường dẫn thì bấm nó: tool ghi lại `Data Source=` cho đúng máy hiện tại, bản cũ lưu thành `Database_Libs.DbLib.<ngày_giờ>.bak` ngay cạnh. Hết cảnh mỗi lần clone repo về là phải mở Notepad sửa tay.

### Đề xuất linh kiện thay thế

Nhập mã cần thay (hoặc mở từ cửa sổ tra cứu) → tool lấy danh sách thay thế của Digi-Key. Nếu Digi-Key không có, tool tự tìm linh kiện **cùng thông số** (giá trị + package + dung sai).

Bảng hiện In-Stock và State từng mã, tick **Chỉ hiện loại còn hàng** để lọc.

Chọn mã xong, tool **luôn đưa bạn sang tab *Thêm linh kiện*** chứ không ghi thẳng — ở đó bạn xem đầy đủ dữ liệu sẽ ghi vào từng cột, đổi **Loại** (sheet) nếu tool đoán sai, sửa tay ô nào cần, bấm **Download datasheet**, rồi mới **Ghi vào Excel**.

Khi chuyển tab, tool **tự tra cứu lại từng mã** bằng đúng nguồn đang chọn ở ô *Nguồn*. Lý do: API danh sách thay thế của Digi-Key chỉ trả về bản rút gọn (thiếu thông số kỹ thuật, mô tả chi tiết, link datasheet), bung thẳng sang tab chính thì bảng *Dữ liệu sẽ ghi* trống nhiều ô. Nhật ký hiện `✔ <mã>: N thông số` cho từng mã; mã nào tra cứu hỏng thì giữ nguyên dữ liệu cũ từ bảng thay thế chứ không mất.

- **nháy đúp một dòng** → mở ngay mã đó ở tab Thêm linh kiện
- tích cột ✓ nhiều mã rồi bấm **Mở ở tab Thêm linh kiện →** → mã đầu tiên hiện luôn ở tab chính, các mã còn lại nằm trong bảng hàng loạt, nháy đúp dòng nào thì mã đó nhảy sang tab chính

### Kiểm tra symbol / footprint

Rà mọi dòng, báo các lỗi khiến DBLib gãy:

- thiếu `Library Path` / `Footprint Path`
- file `.SchLib` / `.PcbLib` không tồn tại
- thiếu `Library Ref` / `Footprint Ref`
- **tên symbol/footprint không có trong file** (đọc trực tiếp nội dung file thư viện)
- **footprint chưa gắn 3D model** — tick **Kiểm tra 3D model** (1.6). Tool đọc từng storage footprint trong `.PcbLib`, tìm bản ghi `ComponentBody` (`MODEL.EMBED` / `MODELID`) nên biết được footprint nào chỉ có 2D. Những tên không có trong file thì đã báo ở mục trên, không đếm trùng vào đây.

Đường dẫn kiểu `.\Symbols\Resistor.SchLib` được tính từ **Gốc thư viện Altium** trong Cấu hình API; để trống thì lấy thư mục chứa file Excel.

Từ 1.5.0 tool có trình đọc file OLE của Altium **viết sẵn trong `altium_ole.py`**, không cần cài `olefile` nữa. Tên symbol lấy từ `FileHeader` (`LibRef0=`, `LibRef1=`...), tên footprint lấy từ bảng tên cuối stream `Library/Data` — đúng nguồn Altium dùng, nên mã có dấu `/` như `USB2514B/M2` không còn bị báo nhầm.

Nếu một file thư viện **không đọc được**, tool báo hẳn một dòng đỏ *"Không đọc được file thư viện — các dòng dùng file này CHƯA được kiểm tra"* và dòng trạng thái luôn ghi *đọc được N/M file thư viện*. Không bao giờ im lặng báo "không có vấn đề" nữa.

### Phát hiện trùng thông số

Tìm các linh kiện **khác mã nhưng cùng bộ thông số** — cùng giá trị, package, dung sai, điện áp. Với sheet ít cột thông số (IC, Connector) thì so thêm cả mô tả. Chống thư viện phình vì mua mỗi lúc một hãng.

### Đối chiếu BOM

Thả file BOM (.csv hoặc .xlsx) xuất từ Altium vào:

- Tự nhận cột part number, số lượng, designator — ưu tiên `Manufacturer Part Number`, cùng lắm mới dùng `Comment`
- Mã nào chưa có trong thư viện tô **đỏ**, mã NRND/Obsolete/LTB cũng đỏ
- **Thêm các mã chưa có vào thư viện** → tra cứu hàng loạt rồi ghi
- Xuất CSV để gửi cho bộ phận mua hàng

### Giá BOM  *(1.6)*

Thả cùng file BOM vào, nhập **số board** cần làm, tool tra giá từng mã rồi tính ra tiền:

| Cột | Ý nghĩa |
|---|---|
| SL/board | lấy từ cột Quantity của BOM, không có thì đếm designator |
| Tổng cần | SL/board × số board |
| SL đặt | làm tròn lên theo **MOQ** và bội số đặt hàng |
| Đơn giá · Bậc | giá theo **bậc số lượng** (1 / 10 / 100 / 1000…) ứng với số lượng đặt |
| Thành tiền | đơn giá × SL đặt |
| In-Stock · Đủ hàng | tồn kho tại thời điểm tra; không đủ thì tô cam, hết hàng tô đỏ |
| MOQ · Lead time | lấy thẳng từ nhà cung cấp |

Dưới cùng là **tổng đơn hàng** và **giá mỗi board**. Điền ô *1 USD = … VND* thì có luôn cột quy đổi (nhận `26500`, `26.500`, `26,500`).

- Tick **Chỉ hiện mã thiếu hàng / không có giá** để soi nhanh chỗ tắc
- Nháy đúp một dòng → nhảy sang tab *Đề xuất thay thế* cho đúng mã đó
- **Tính lại (không tra cứu)** khi chỉ đổi số board — không gọi API lần nữa
- Xuất CSV gửi mua hàng, có đủ số lượng đặt, đơn giá, thành tiền, quy đổi

Giá lấy từ Digi-Key (`ProductVariations[].StandardPricing`) và Mouser (`PriceBreaks`). Nếu sheet của bạn có sẵn cột tên `Unit Price` / `Giá` / `MOQ` / `Lead time` thì lúc thêm linh kiện tool điền luôn vào đó.

> Giá và tồn kho là ảnh chụp lúc tra cứu, dùng để so sánh và ước lượng — không phải báo giá chính thức.

### Where-used  *(1.6)*

Trỏ tới thư mục chứa các project Altium, tool đọc mọi file `.SchDoc` (nhóm theo `.PrjPcb`) và cho biết **mã nào đang nằm trên board nào**:

| Cột | Ý nghĩa |
|---|---|
| Design Item ID | mã đọc được từ schematic |
| Trong thư viện | có trong `Master_Library.xlsx` hay không |
| State | trạng thái đang lưu trong thư viện |
| Số lần dùng · Số project · Project · Designator | R12, C7… lấy đúng theo `OWNERINDEX` của Altium |

- Tick **Chỉ hiện mã có vấn đề** → còn lại các mã chưa có trong thư viện hoặc đang NRND / Obsolete / Last Time Buy. Đây chính là danh sách "board nào bị ảnh hưởng khi mã này chết".
- Nút **Mã trong thư viện không board nào dùng** → soi chiều ngược lại
- Chọn một mã, bảng dưới liệt kê từng file dùng nó; **Mở thư mục file đang chọn** để nhảy tới ngay
- Xuất CSV được

Cách đọc: `.SchDoc` là file OLE, các bản ghi nằm trong stream `FileHeader` dạng `|RECORD=1|LIBREFERENCE=…|DESIGNITEMID=…`. Tool tự tách bản ghi, ghép designator (`RECORD=34`) vào linh kiện cha qua `OWNERINDEX`; file lạ thì quay về quét chuỗi thô nên vẫn ra được danh sách mã.

### Thống kê  *(1.6)*

Một trang tổng quan: tổng số linh kiện, tỉ lệ có datasheet / có Manufacturer, số dòng thiếu `Footprint Ref`, số nhóm trùng thông số; biểu đồ cột **theo nhóm** và **theo State** (vẽ thẳng bằng Canvas, không cần cài thư viện đồ hoạ); bảng nhà sản xuất và package dùng nhiều nhất.

**Xuất báo cáo HTML** ra `<Gốc xuất>\thong_ke\thong_ke_<ngày_giờ>.html` rồi mở bằng trình duyệt — tiện gửi cho người khác xem. Xuất CSV cũng được.

## 3e. Datasheet ngay ở tab chính

Tra cứu xong, nút **Download datasheet** đổi màu theo kết quả:

| Nút | Nghĩa |
|---|---|
| **Download datasheet ✓** (xanh lá) | tìm được link, bấm để tải ngay |
| **Không có datasheet ✗** (cam đỏ) | nhà cung cấp không trả về link |

Tick **Tải datasheet khi ghi** thì mỗi lần bấm *Ghi vào Excel* datasheet tự tải luôn ở luồng nền, không phải chờ.

Trong **Thêm hàng loạt** có thêm cột **DS**: nháy vào ô đó để chọn mã nào cần tải, `—` nghĩa là mã đó không có link. Kèm nút **Tick DS tất cả** và **Tải datasheet đã tick**. Khi bấm *Ghi các mã đã chọn*, mọi mã có tick DS được tải xuống sau khi ghi xong.

Thư mục lưu đặt trong **Cấu hình → Thư mục → Lưu datasheet**.

### Xem datasheet  *(1.6.3)*

Cạnh nút tải là nút **Xem datasheet**:

- Đã tải file về máy rồi → nút ghi **Xem datasheet (máy)**, bấm là mở thẳng file PDF bằng trình đọc mặc định — nhanh và không cần mạng
- Chưa tải → **Xem datasheet (online)**, mở link do API trả về bằng trình duyệt mặc định
- Link kiểu `//mm.digikey.com/...` (thiếu `https:`) được tự bù trước khi mở

Tab **Tra cứu thư viện** cũng có nút này: chọn một mã rồi bấm là xem ngay; dòng nào chưa có link trong Excel thì tool tra cứu online rồi mở.

Không thích thì bỏ tick **Cấu hình → Giao diện → Hiện nút "Xem datasheet"**, nút sẽ ẩn hẳn.

## 3f. Giao diện co giãn

Cửa sổ tự sắp lại theo kích thước — nửa màn hình dọc, nửa ngang, hay toàn màn hình đều dùng được:

- Thanh tab tự xuống nhiều hàng khi cửa sổ hẹp
- Ô tra cứu gộp một hàng khi rộng, tách hai hàng khi hẹp
- Bảng *Dữ liệu sẽ ghi* và *Nhật ký* nằm trong khung kéo thả được — kéo đường phân cách để chia lại chỗ
- Mọi bảng và ô nhật ký đều có **thanh cuộn dọc và ngang**

**Chế độ tối (1.6):** nút **🌙 Tối / ☀ Sáng** ngay trên thanh công cụ, đổi là ăn ngay cả thanh tab, bảng, ô nhập lẫn cửa sổ phụ. Chọn cố định `light` / `dark` / `auto` (theo cài đặt Windows) ở **Cấu hình → Giao diện**. Mọi màu trong app lấy từ một bảng màu chung trong `theme.py` nên không còn cảnh chữ đen trên nền đen.

**Ghim tab (1.6.1):** 12 tab nhét vào một hàng thì tên nào cũng bị cắt còn `Thêm lin...`, `Tra cứu thư ...`. Bấm nút **Tab...** trên thanh công cụ (hoặc **Cấu hình → Giao diện → Ghim tab...**) rồi bỏ tick những tab ít dùng — các tab còn lại tự rộng ra, đọc được tên đầy đủ.

- Nút **Bộ gọn** để sẵn 6 tab hay dùng nhất: Thêm linh kiện · Tra cứu · Quét State · Sức khoẻ · Đối chiếu BOM · Liên hệ
- Tab *Thêm linh kiện* luôn hiện, không bỏ được
- **Tab bị ẩn không mất đi**: mở nó từ chỗ khác (ví dụ bấm *Đề xuất thay thế* trong tab Tra cứu) thì nó tự hiện lại tạm thời cho tới lần bạn đổi lựa chọn ghim
- Lựa chọn lưu trong `config.json` ở khoá `visible_tabs`

**Rê chuột vào tab (1.6.1):** giữ chuột trên một tab khoảng nửa giây sẽ hiện thẻ nhỏ ghi **tên đầy đủ** và một dòng mô tả tab đó làm gì — kiểu thẻ xem trước của trình duyệt. Thẻ hiện **ngay dưới tab** đang trỏ. Tắt được, và chỉnh được độ trễ, ở **Cấu hình → Giao diện** (1.6.3).

### Chọn logo  *(1.6.9)*

**Functions → Giao diện → Logo**: hai lựa chọn, có ảnh xem trước ngay cạnh mỗi lựa chọn.

| Lựa chọn | File | Ghi chú |
|---|---|---|
| **Ếch** | `frog.png` · `frog.ico` · `frog_icon.png` | mặc định từ 1.6.9 |
| **Nhện** | `icon.png` · `icon.ico` | logo cũ, vẫn giữ nguyên trong bản cài |

Ảnh lớn ở tab Liên hệ và ảnh icon 16 px không thể là một: ảnh lớn nhiều nền trắng, thu về 16 px chỉ còn một ô trắng. Nên từ **1.7.1**, cạnh `<tên>.png` mà có file `<tên>_icon.png` — bản cắt sát vào chủ thể — thì icon cửa sổ dùng file đó. Không có thì dùng ảnh lớn như cũ.

Đổi là **thấy ngay, không cần mở lại tool**: cả ảnh lớn ở tab **Liên hệ** lẫn icon trên **thanh tiêu đề** và thanh taskbar đều đổi theo. Lựa chọn lưu ở khoá `logo` trong `config.json` (`"frog"` hoặc `"spider"`).

Ảnh thu nhỏ bằng **bội số nguyên** tính từ bề ngang thật của file, nên thay `frog.png` bằng ảnh khác kích thước vẫn ra cỡ tương đương — không phải sửa code. Thiếu file ảnh thì tool tự lùi về logo còn lại chứ không vỡ giao diện.

## 3h. Cache tra cứu  *(1.6)*

Kết quả tra cứu được cất vào SQLite (`lookup_cache.db`) theo khoá *(kiểu tra cứu, nguồn, part number)*. Tra lại một mã đã tra gần đây thì lấy thẳng từ cache — nhật ký ghi `(cache) <mã>`.

Chỉnh trong **Cấu hình → Cache · Giao diện · Cập nhật**:

| Ô | Mặc định | Ý nghĩa |
|---|---|---|
| Bật cache | có | tắt là quay về hành vi 1.5.x, lần nào cũng gọi API |
| Giữ kết quả trong … giờ | 12 | quá hạn thì coi như chưa có, gọi API lại |
| Tối đa … mục | 20.000 | vượt thì tự xoá các mục cũ nhất |

Kèm nút **Xoá toàn bộ cache** và **Dọn mục quá hạn**, dòng trạng thái cho biết đang có bao nhiêu mục, bao nhiêu MB, phiên này trúng / trượt bao nhiêu lần.

Vì sao đáng làm: quét State cả thư viện 214 mã là 214 lần gọi API tuần tự cộng thời gian nghỉ giữa các lần — vài phút và ăn vào hạn mức API. Có cache thì lần quét thứ hai trong ngày gần như tức thì, và lúc lấy từ cache tool cũng không nghỉ nữa.

## 3i. Cập nhật tool  *(1.6)*

Tab **Liên hệ** có khung Cập nhật: **Kiểm tra bản mới** → **Tải & cập nhật** → **Hoàn tác bản cập nhật cuối**. Từ **1.6.6** không cần vào tab đó nữa: mở tool là tự kiểm tra, có bản mới thì hiện hộp thoại hỏi luôn.

### Hộp thoại "Có bản mới"  *(1.6.6)*

Mở tool khoảng 1,5 giây sau, nếu trên GitHub có bản cao hơn bản đang chạy thì hiện hộp thoại ghi số phiên bản, nguồn tìm thấy, ngày phát hành và phần **Có gì mới** lấy từ ghi chú phát hành. Ba lựa chọn:

| Nút | Việc xảy ra |
|---|---|
| **Cập nhật ngay** | Tải `.zip` và cài luôn, không hỏi lại lần nữa |
| **Bỏ qua phiên bản này** | Ghi số phiên bản đó vào `update.skip_version`; **mở tool lần sau không hỏi lại bản này nữa**, nhưng có bản cao hơn thì vẫn hỏi. Bấm **Kiểm tra bản mới** thủ công thì vẫn hiện |
| **Đóng** (hoặc Esc / dấu ✕) | Bỏ qua lần này thôi, lần sau mở tool vẫn hỏi |

Muốn hỏi lại bản đã bỏ qua: **Cấu hình → Cập nhật → Hỏi lại bản đã bỏ qua**. Không muốn bị hỏi lúc mở tool thì bỏ tick *Kiểm tra bản mới khi mở tool*.

### Chỉ một nguồn: GitHub Releases  *(1.6.7)*

Tool **chỉ** tìm bản mới ở `api.github.com/repos/<repo>/releases/latest`. Bản phát hành phải **đính kèm một file `.zip`** (asset) thì mới tải tự động được; Release không có file `.zip` thì tool nói thẳng chứ không im lặng.

Cách phát hành một bản mới: GitHub → **Releases → Draft a new release** → tag đặt kiểu `v1.7.0` → kéo file `AltiumDbTool_v1.7.0.zip` vào phần assets → Publish.

> ⚠ **Thả file vào đúng chỗ.** Trang tạo Release có **hai** vùng nhận file, nhìn rất giống nhau:
>
> | Vùng | Kết quả | Tool tự cập nhật được? |
> |---|---|---|
> | Ô soạn **ghi chú phát hành** (*Describe this release*) | GitHub tải file lên dạng *user-attachment* rồi **chèn một đường link** vào phần mô tả. Trường `assets` của API vẫn **rỗng** | Từ **1.7.0** thì được — trước đó thì không |
> | Khung **Assets** — *"Attach binaries by dropping them here"*, nằm **dưới** ô ghi chú | File thành **asset** thật của Release | Có, mọi phiên bản |
>
> Trang web hiển thị cả hai gần như giống nhau nên rất dễ nhầm. Thả nhầm vào ô ghi chú là bản 1.6.8 báo *"không đính kèm file .zip nào"* dù bạn nhìn thấy file trên trang — tool đọc `assets` của API, mà chỗ đó rỗng thật.
>
> Từ **1.7.0** tool vớt nốt trường hợp này: không thấy asset thì nó dò tiếp **link `.zip` trong phần ghi chú**. Chỉ nhận link trỏ tới tên miền của GitHub (`github.com`, `objects.githubusercontent.com`, `raw.githubusercontent.com`…), để một dòng chữ trong ghi chú không dắt tool đi tải file ở đâu khác về chạy. Có nhiều link thì ưu tiên link có số phiên bản trong tên. Asset thật vẫn được ưu tiên hơn link trong ghi chú.

> **Không dùng "Source code (zip)"** của GitHub. Đó là ảnh chụp **mã nguồn của nhánh** tại lúc đặt tag, không phải bản đóng gói. Số phiên bản thì lấy theo **tag**, còn nội dung lại lấy theo **mã nguồn** — hai thứ này lệch nhau ngay khi bạn đặt tag `V1.6.8` mà chưa push code 1.6.8 lên nhánh. Từ **1.6.8** tool bỏ hẳn đường lui về đây; trước đó Release thiếu asset là tool âm thầm tải mã nguồn về chép đè, báo "đã cập nhật lên 1.6.8" trong khi thực tế **hạ cấp** máy về đúng bản đang nằm trên nhánh.

### Ba lớp kiểm tra trước khi chép đè  *(1.6.8)*

Tải xong tool **soi file rồi mới chép**, sai một trong ba là dừng luôn và **không đụng gì vào máy**:

1. File có mở được không, có đúng là bản đóng gói của tool không
2. Số phiên bản **bên trong file** (`APP_VERSION`) có khớp với số ghi trên Release không — lệch là dấu hiệu Release đính kèm nhầm file, hoặc code chưa push lên cùng tag
3. File có **cũ hơn** bản đang chạy không — cũ hơn là chặn, tránh vô tình hạ cấp

Chép xong tool còn **đọc lại `altium_db_tool.py` trên đĩa** để chắc phiên bản đúng như mong đợi. Không đúng thì báo ngay kèm đường dẫn thư mục `_update_backup` để bấm **Hoàn tác bản cập nhật cuối**. Vẫn muốn cài đúng file đó thì tải về rồi dùng nút **Cập nhật từ file .zip...** — đường cài tay không bị ba lớp này chặn, chỉ hỏi lại cho chắc.

Trước 1.6.7 tool còn dò thêm thư mục `.zip` trong repo và file `version.json`. Hai đường đó nhặt **bất kỳ** file `.zip` nào có số phiên bản trong tên, nên dễ bắt nhầm bản cũ hay bản thử nghiệm nằm chung repo — đã bỏ. Bản lẻ thì cài tay, xem mục dưới.

### Cập nhật thủ công từ file `.zip`  *(1.6.7)*

Nút **Cập nhật từ file .zip...** — có ở cả tab **Liên hệ** lẫn **Cấu hình → Cập nhật**. Dùng khi repo chưa tạo Release nào, khi nhận bản qua USB / chat, hoặc khi muốn quay lại một bản cũ trong thư mục `History\`.

Chọn file xong tool **soi file trước khi cài**: kiểm tra zip có hỏng không, có đúng là bản của tool không (phải chứa `altium_db_tool.py`), đọc số phiên bản từ dòng `APP_VERSION` bên trong file — chứ không tin vào tên file — rồi hỏi lại kèm số phiên bản và số file. Cài bản **cũ hơn** hoặc **cùng phiên bản** với bản đang chạy vẫn được, nhưng có dòng cảnh báo trước.

File `.zip` bạn chọn **không bị xoá** sau khi cài. Phần còn lại giống hệt cập nhật tự động: `config.json`, `rules.json`, `api_keys.enc`, cache và nhật ký giữ nguyên, bản cũ cất trong `_update_backup\` nên **Hoàn tác bản cập nhật cuối** vẫn dùng được.

### Nguồn cập nhật bị ẩn, mở bằng mật khẩu  *(1.7.6)*

Từ 1.7.6, hai ô **Repo GitHub** và **GitHub token** trong *Cấu hình → Cập nhật*
không hiện sẵn nữa. Chỗ đó chỉ còn một dòng **🔒 Đã ẩn — cần mật khẩu để xem và
sửa** kèm nút **Mở khoá...**; nhập đúng mật khẩu thì hai ô hiện ra như cũ, và có
thêm nút **Ẩn lại**. Mở khoá **chỉ có hiệu lực tới khi đóng tool** — không ghi vào
`config.json`, nên lần mở sau lại khoá như cũ.

Lý do: hai ô này trỏ tới kho phát hành của tool. Sửa nhầm vào đây là tool im lặng
không thấy bản mới nào nữa, mà lỗi thì không nhìn ra ngay.

**Ẩn thật, không chỉ ẩn cái ô.** Khoá một ô nhập mà câu báo lỗi vẫn in nguyên
đường dẫn ra thì khoá để làm gì. Nên khi đang khoá, tên repo được thay bằng
`(đã ẩn)` ở **mọi chỗ nó có thể lọt ra**:

- các câu lỗi của `check()` (404, token sai, chưa có Release…)
- dòng ghi trong ô **Nhật ký**
- **lỗi mạng của `requests`** — chỗ dễ sót nhất: nó ném kèm nguyên URL, kiểu
  `Max retries exceeded with url: /repos/<owner>/<tên-repo>/releases/latest`.
  Chỉ cần rút mạng rồi bấm *Kiểm tra ngay* là cả đường dẫn hiện ra trong hộp
  thoại. Hàm `updater.scrub()` xoá cả ba dạng: `owner/tên`, `owner%2Ftên`
  (URL-encode) và riêng phần `tên`.

Nút **Kiểm tra ngay** và **Cập nhật từ file .zip...** vẫn dùng được bình thường
khi đang khoá — chỉ có việc *xem và sửa địa chỉ* là cần mật khẩu.

> **Đây là cái chốt cửa, không phải ổ khoá.** Mật khẩu nằm trong chính chương
> trình, nên ai quyết tâm vẫn lấy ra được — đọc mã nguồn, hoặc giải nén file
> `.exe` rồi đọc `.pyc`. Cất ở dạng băm (`sha256` kèm muối, hằng số
> `CFG_PASS_HASH` trong `updater.py`) chỉ chặn được đường dễ nhất là mở file ra
> thấy ngay chuỗi mật khẩu; mật khẩu bốn chữ số thì dò hết 10.000 khả năng chỉ
> mất một nháy mắt. Mục đích là **ngăn sửa nhầm**, không phải giữ bí mật — đừng
> đặt thứ gì thật sự cần giữ kín sau cái chốt này.
>
> Đổi mật khẩu: chạy dòng dưới rồi dán kết quả vào `CFG_PASS_HASH` trong
> `updater.py`.
>
> ```
> python -c "import hashlib;print(hashlib.sha256(('AltiumDbTool.v1:'+'matkhaumoi').encode()).hexdigest())"
> ```

### Ô "Repo GitHub" nhận cả đường dẫn đầy đủ  *(1.6.6)*

Dán `https://github.com/Nhatkhongbuon/Altium-DB-Library-Helper` vào ô Repo là **mọi** lời gọi GitHub trở thành `api.github.com/repos/https:/github.com/...` nên trả về 404 hết — tool báo "không liên lạc được GitHub / chưa có bản phát hành nào" dù repo vẫn sống. Đây chính là lý do tính năng cập nhật không chạy trước 1.6.6.

Nay ô này nhận mọi cách viết và tự cắt về dạng `owner/repo`:

| Gõ vào | Tool hiểu là |
|---|---|
| `https://github.com/Nhatkhongbuon/Altium-DB-Library-Helper` | `Nhatkhongbuon/Altium-DB-Library-Helper` |
| `https://github.com/…/Altium-DB-Library-Helper.git` | như trên |
| `git@github.com:Nhatkhongbuon/Altium-DB-Library-Helper.git` | như trên |
| `.../tree/main/History` | `Nhatkhongbuon/Altium-DB-Library-Helper` |

Cấu hình cũ có sẵn URL cũng được dọn **tự động lúc mở tool**, không phải sửa tay. Bên cạnh còn nút **Kiểm tra ngay** — lưu ô repo/token rồi kiểm tra luôn, khỏi phải sang tab Liên hệ.

### Repo private thì phải có GitHub token  *(1.6.2)*

Repo **public** thì không cần làm gì. Repo để **private** mà gọi ẩn danh thì GitHub trả về **404** — y hệt như repo không tồn tại — nên tool sẽ báo "không có bản phát hành nào" dù thực ra có.

Cách xử lý: **Cấu hình → Cập nhật → GitHub token**.

1. github.com → **Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token**
2. **Repository access**: chọn đúng repo thư viện này
3. **Repository permissions → Contents: Read-only** (chỉ cần thế, không cần quyền gì khác)
4. Dán token vào ô, bấm Lưu

Token được **mã hoá chung với API key** trong `api_keys.enc`, không nằm trong `config.json`. Có token rồi thì cả ba đường tìm bản mới đều đi qua GitHub API kèm `Authorization`, và file `.zip` của release được tải bằng URL API của asset (`Accept: application/octet-stream`) vì `browser_download_url` không tải ẩn danh được trên repo private.

Nếu chưa nhập token mà repo private, phần Cập nhật báo thẳng: *"GitHub trả về 404. Repo đang PRIVATE thì gọi ẩn danh luôn bị báo 404 như vậy — nhập GitHub token ở Cấu hình → Cập nhật."*

Cài đặt: tải `.zip`, giải nén, **sao lưu các file sẽ bị thay** vào `_update_backup\<ngày_giờ>\` rồi mới chép đè. `config.json`, `rules.json`, `api_keys.enc`, cache và nhật ký **không bị đụng tới**. Cài xong đóng và mở lại tool. Thấy bản mới hỏng thì bấm **Hoàn tác bản cập nhật cuối**.

Repo và ô *Kiểm tra bản mới khi mở tool* đặt trong **Cấu hình → Cập nhật**.

## 3j. Tạo file thư viện mới ngay trong tool  *(1.6.3)*

Nút **Tạo mới...** cạnh nút Chọn. Trước đây muốn có file Excel đúng cấu trúc thì phải xin file của người khác rồi xoá dữ liệu đi; giờ tool tự dựng được.

Hộp thoại cho chọn:

| Mục | Ý nghĩa |
|---|---|
| **Bộ mẫu chuẩn của tool** | 13 nhóm sẵn có: Resistor, Capacitor, Polarized Capacitor, Ferrite, Inductor, IC, Diode, TransFet, Oscillator, Connector, Switch, Fuse, Other |
| **Sao cấu trúc của file đang mở** | Lấy đúng cột, công thức và cột cố định của thư viện hiện tại nhưng **không chép dữ liệu** — hợp khi cả team cần một file trống giống hệt file đang dùng |
| **Tự đặt nhóm và trường** *(1.6.4)* | Tự gõ tên nhóm và chọn từng trường — xem ngay bên dưới |
| Tick từng nhóm | Chỉ tạo những nhóm bạn cần |
| **Tạo luôn file .DbLib đi kèm** | Sinh `Database_Libs.DbLib` đã trỏ đúng file Excel, bật sẵn mọi bảng và **map đầy đủ trường** (Design Item ID là Key Field, Library Path/Ref và Footprint Path/Ref là tham số hệ thống, các cột thông số tick VisibleOnAdd) |
| **Dùng file mới này làm thư viện đang mở** | Nạp luôn sau khi tạo |

#### Chế độ tự đặt nhóm và trường  *(1.6.4)*

Cột trái là danh sách nhóm của bạn (gõ tên → **Thêm**; có sẵn Resistor, Capacitor, IC để sửa hoặc xoá). Chọn một nhóm thì cột phải hiện các trường của nó:

- **Bắt buộc** — sheet nào cũng phải có, không bỏ được: `Design Item ID`, `State`, `Library Path`, `Library Ref`, `Footprint Path`, `Footprint Ref`
- **Trường nên có** — tick sẵn, bỏ được: `Package`, `Comment`, `Manufacturer`, `Description`, `Detailed Description`, `Manufacturer Product Number`, `ComponentLink1URL`, `ComponentLink1Description`
- **Trường tự thêm** — gõ tên rồi **Thêm**, hoặc chọn từ ô *gợi ý nhanh* (Tolerance, Power, Voltage - Rated, Frequency, Unit Price…)
- **Symbol dùng chung** — điền tên (ví dụ `Resistor_Standard`) thì cả nhóm dùng chung một symbol; bỏ trống thì symbol đặt tên trùng mã linh kiện
- **File .SchLib / .PcbLib** — bỏ trống thì lấy theo tên nhóm

Công thức được sinh tự động cho nhóm tự đặt: `Manufacturer Product Number` và `ComponentLink1Description` lấy theo `Design Item ID`; `Footprint Ref` lấy theo `Package` nếu nhóm có cột Package, không thì theo `Design Item ID`.

Các ô nhập đều có **chữ gợi ý mờ** (ví dụ *"ví dụ: Relay"*): gõ vào là gợi ý biến mất, xoá hết thì hiện lại, và gợi ý không bao giờ bị tính nhầm thành dữ liệu thật.

File sinh ra có:

- **cột chữ để định dạng Text** nên `Package` `0402` không bị Excel đổi thành số `402` *(1.6.4)*
- mỗi sheet là một **Excel Table** nên thêm dòng là công thức tự lan xuống
- **dòng "mầm" ở hàng 2** giữ sẵn công thức và các cột cố định (`Library Path`, `Footprint Path`, `Library Ref`, `State`) — linh kiện đầu tiên bạn thêm sẽ ghi đè đúng lên dòng đó, **đừng xoá nó đi**
- công thức `Value Sort` bọc `IFERROR` nên dòng trống không hiện `#VALUE!`
- sheet `Other` **có cột State** (bản cũ thiếu)

Sau khi tạo, còn phải tự tạo hai thư mục `Symbols\` và `Footprints\` cạnh file Excel rồi vẽ symbol / footprint vào đó — tool chỉ dựng phần dữ liệu, không vẽ hộ được.

## 3k. Hai kiểu file database  *(1.7.3, đặt lại mặc định ở 1.7.4)*

Tool làm việc được với hai kiểu file. **`Master_Library.xlsx` (From Nhatnpm) là kiểu mặc định** — kiểu tool dùng từ đầu, mỗi sheet một loại linh kiện, cột khoá tên là **`Design Item ID`**. **`Database_Library.xlsx` (From Hoangnh277) là kiểu tuỳ chọn** — kiểu hay gặp khi lấy file mẫu của Altium hoặc nhận thư viện từ nơi khác.

"Mặc định" ở đây chỉ là **thứ tự và cách gọi tên khi hiện ra màn hình**: kiểu mặc định đứng trước, được chọn sẵn trong hộp chuyển đổi, và tên kiểu ghi kèm tác giả cho rõ nguồn gốc. Về đọc / ghi / chuyển đổi thì **hai kiểu chạy y hệt nhau** — mở file nào tool cũng tự nhận ra kiểu của nó, không phải khai báo gì.

Trước 1.7.3 tool chỉ hiểu đúng kiểu mặc định; mở file kiểu kia thì nạp được sheet nhưng không tìm thấy cột khoá, nên bảng thông số trống trơn, tra trùng không chạy, và ghi vào là hỏng dòng.

| | Master_Library *(From Nhatnpm — mặc định)* | Database_Library *(From Hoangnh277 — tuỳ chọn)* |
|---|---|---|
| Mỗi sheet là | một **loại** linh kiện (Resistor, Capacitor, IC…) | một **nhóm lớn** (PASSIVE, SEMICONDUCTOR, IC, CONNECTOR, ELECTROMECHANICAL) |
| Cột khoá | `Design Item ID` | `Mfr. #` |
| Dữ liệu nằm trong | ô Excel thường (có thể có Table) | **Excel Table** (ListObject) — bắt buộc |
| Loại linh kiện ghi ở | tên sheet | cột **`Type`** trong sheet |
| SMD hay xuyên lỗ | cột `Mounting Type` | cột **`Footprint type`** (`SMD` / `TH`) |
| Kích thước / vỏ | `Package` | **`Size`** |
| Nhà sản xuất | `Manufacturer` | **`Manufacturer 1`** |
| Vòng đời | `State` | không có |
| Cột thêm | `Value Sort`, `ComponentLink1*` | `Footprint Ref 2` / `Footprint Path 2`, `Pick and Place` |

**Không phải chọn kiểu ở đâu cả.** Bấm **Chọn...** rồi trỏ tới file nào cũng được — tool đọc cấu trúc rồi tự biết, và ghi một dòng vào Nhật ký:

```
Đã nạp 5 sheet: PASSIVE, SEMICONDUCTOR, IC, CONNECTOR, ELECTROMECHANICAL
  • Định dạng: Database_Library (From Hoangnh277) — kiểu tuỳ chọn, bảng Excel Table, khoá “Mfr. #”
```

Mở file kiểu mặc định thì dòng đó là:

```
  • Định dạng: Master_Library (From Nhatnpm) — kiểu mặc định, khoá “Design Item ID”
```

Mọi công cụ có sẵn — Tra cứu thư viện, Kiểm tra thư viện, Symbol/Footprint, Trùng thông số, Where-used, Thống kê, Đối chiếu BOM, Backup — chạy nguyên vẹn trên cả hai kiểu; chỗ nào trước ghi *"Design Item ID"* thì với file kiểu mới nó hiểu là *"Mfr. #"*.

### Ghi thêm dòng vào Excel Table

Dữ liệu của `Database_Library.xlsx` nằm trong Excel Table. Thêm một dòng ngay dưới bảng mà **không nới phạm vi bảng** thì Excel lẫn Altium đều coi như dòng đó không tồn tại — nhìn trong Excel thấy chữ nhưng query ra không có. Nên khi ghi bằng `openpyxl` (máy không cài Excel), tool sửa cả `ref` của Table lẫn `ref` của bộ lọc:

```
Table1  A1:S328   →   A1:S329
```

Ghi **đè** lên một dòng đã có thì bảng giữ nguyên kích thước, không phình thêm.

### Công thức `=Table1[[#This Row],[Mfr. '#]]`

Kiểu file này dùng **structured reference** — công thức trỏ theo *tên cột* chứ không theo toạ độ ô, và tự hiểu "dòng của chính nó". Chép xuống dòng mới thì phải giữ **nguyên văn**: dịch toạ độ như công thức thường là hỏng. Tool nhận ra dạng này và chép nguyên. Dấu `'` trong `Mfr. '#` là ký tự thoát của Excel cho dấu `#`, tool cũng hiểu.

### Công thức trỏ sang file Excel khác → tool điền giá trị thật

File `Database_Library.xlsx` đi vòng qua nhiều máy thường còn sót vài ô kiểu:

```
=TRIM(MID([1]!Table33[[#This Row],[Description]], ...))
```

`[1]!` là **một file Excel khác** — trên máy bạn thường là `#REF!`. Chép công thức đó xuống dòng mới chỉ tạo thêm `#REF!`.

Nên từ 1.7.3 tool chỉ coi một cột là **"cột công thức"** (dòng xanh, không ghi đè) khi:

- công thức xuất hiện ở **phần lớn** (≥ 60%) các dòng đã có dữ liệu, **và**
- công thức **không** trỏ sang file Excel khác.

Không đạt thì cột đó được coi là cột dữ liệu bình thường — tool điền **giá trị thật** lấy từ API vào. Lúc nạp file, Nhật ký nói rõ cột nào rơi vào diện này:

```
  • Cột có công thức lẻ / trỏ sang file Excel khác — tool điền giá trị thật
    thay vì chép công thức: Footprint Ref, Library Ref, Size, Value, Voltage
```

Tab **Kiểm tra thư viện** cũng có thêm mục **"Ô đang báo lỗi Excel"** — đếm các ô `#REF!`, `#N/A`, `#VALUE!`… theo từng cột, kèm ví dụ dòng đầu tiên, để bạn biết chỗ nào cần gõ lại tay.

### Đoán sheet cho file gộp nhóm

`rules.json` của bạn có từ khoá cho các sheet kiểu `Resistor`, `Capacitor`… nhưng file kiểu mới lại đặt tên sheet là `PASSIVE`, `IC`… Không khớp từ khoá nào thì tool dùng **bảng từ khoá mặc định theo nhóm**: điện trở / tụ / cuộn cảm / ferrite / thạch anh → `PASSIVE`, diode / MOSFET / LED / opto → `SEMICONDUCTOR`, MCU / FPGA / regulator / transceiver → `IC`, header / terminal block / jack → `CONNECTOR`, nút bấm / relay / buzzer / test point → `ELECTROMECHANICAL`.

### `Footprint Ref` không đoán bừa

Ở `Master_Library.xlsx` mỗi sheet chỉ có một loại linh kiện nên `0402 → R0402` là chắc chắn. Ở sheet `PASSIVE` gộp cả điện trở, tụ, cuộn cảm và ferrite thì riêng `0402` **không nói lên điều gì** — đúng footprint nào còn tuỳ loại.

Nên với sheet có cột `Type`, tool chỉ điền `Footprint Ref` khi khớp **cả `Type` lẫn `Size`** (`RESISTOR` + `0805` → `RESC 0805(2012) SMD`), hoặc khi kích thước đó trong cả sheet chỉ ứng với **đúng một** footprint (`DO-214AB`). Không chắc thì để trống — ô đỏ nhắc bạn tự điền, hơn là điền nhầm footprint của linh kiện khác. Mục kiểm tra *"Footprint Ref khác quy luật của sheet"* cũng tính theo cặp `Type / Size` nên hết báo nhầm hàng loạt.

### Chuyển đổi qua lại

**Functions → Thư viện → Chuyển đổi định dạng DB**. Trang này đọc file đang mở, cho biết nó đang là kiểu gì, và cho chuyển sang kiểu còn lại:

- Bảng **xem trước** liệt kê sheet nguồn nào chạy về sheet đích nào, kèm số dòng — xem xong mới bấm chuyển.
- Kết quả **luôn ghi ra một file mới**, tool không đụng vào file gốc. Chọn trùng tên file nguồn là bị từ chối.
- Chuyển sang `Database_Library`: các sheet được gom về 5 nhóm lớn theo tên sheet (nhóm nào không đoán ra thì vào sheet `OTHER`), mỗi sheet đích được tạo sẵn một **Excel Table**.
- Chuyển sang `Master_Library`: tách sheet theo cột `Type`; các cách viết khác nhau của cùng một loại (`CAPACITOR-ALUMIUM`, `Capacitor Alumium`) được gom làm một, tên sheet lấy cách viết phổ biến nhất trong file. Thêm sẵn cột `State` = `Active`.
- Tick **Giữ lại các cột lạ** (mặc định bật): cột nào bên kia không có chỗ nhận thì được đẩy về cuối sheet chứ không bị bỏ đi.
- Xong tool hỏi *"Mở file mới này trong tool luôn?"*.

Chỉ chép **giá trị đang hiện**, không chép công thức — công thức trỏ theo toạ độ cũ hay trỏ sang file khác mà mang nguyên sang file mới thì thành `#REF!` hết. Sang file mới bạn dựng lại công thức cho cột nào muốn Excel tự tính.

### File `.DbLib`

Khi tool sinh `.DbLib` (mục **3j**), cột khoá được map thành tham số **`Design Item ID` với `FieldType=0`** dù tên cột trong Excel là gì — Altium cần đúng cái tên tham số đó để coi là khoá chính, nhưng tên cột thì đặt gì cũng được. File kiểu mới vì thế map `Mfr. #` → `Design Item ID`.

## 3l. Bôi đen & copy trong bảng  *(1.7.4)*

Mọi bảng trong tool — kết quả tra cứu, Sức khoẻ thư viện, Symbol / Footprint, Trùng thông số, Where-used, Thống kê, Đối chiếu BOM, danh sách backup, History, bảng xem trước khi chuyển đổi — giờ **bôi đen và copy được**.

Trước đây không copy được là do bản chất của widget bảng: nó **vẽ chữ lên canvas** chứ không phải ô văn bản, nên kéo chuột không chọn được ký tự nào và Ctrl+C không làm gì cả. Nhìn thấy `ADS8688AIDBT` báo thiếu symbol ngay đó mà muốn dán sang Altium thì phải gõ lại tay — gõ lại là sai chính tả, mà sai một ký tự thì tra ra rỗng.

| Thao tác | Kết quả |
|---|---|
| **Ctrl+C** | Chép các dòng đang chọn, các cột cách nhau bằng **Tab** — dán thẳng vào Excel là mỗi cột vào đúng một ô |
| **Ctrl+Shift+C** | Như trên nhưng **kèm dòng tiêu đề** |
| **Ctrl+A** | Chọn tất cả các dòng |
| Kéo chuột / **Shift+click** / **Ctrl+click** | Chọn nhiều dòng liền nhau hoặc rời nhau |
| **Nháy phải** | Menu: *Sao chép ô · Sao chép dòng đang chọn · Sao chép dòng kèm tiêu đề · Sao chép cột · Sao chép cả bảng · Bôi đen chữ trong ô… · Chọn tất cả* |
| **Enter** | Mở lớp phủ **bôi đen chữ ngay trong ô** đang chọn |
| **Nháy đúp** | Cũng mở lớp phủ bôi đen — trừ những bảng vốn đã dùng nháy đúp cho việc khác (sửa ô ở tab *Thêm linh kiện*, xem chi tiết ở *Tra cứu thư viện*…); ở các bảng đó dùng **Enter** hoặc menu chuột phải |

### Lớp phủ bôi đen trong ô

Bấm **Enter** (hoặc nháy phải → *Bôi đen chữ trong ô…*) thì một ô nhập **chỉ đọc** phủ đúng lên ô đang chọn, chữ được chọn sẵn:

- kéo chuột chọn **một đoạn** bất kỳ — ví dụ chỉ lấy `RJHSE5380` trong cả câu báo lỗi;
- **Ctrl+C** chép phần đã chọn;
- ô dài hơn bề rộng cột thì dùng mũi tên trái / phải để chạy hết chữ, không phải kéo giãn cột;
- **Tab / Shift+Tab** sang ô bên cạnh, **↑ / ↓** sang dòng trên dưới — soi cả bảng không cần đóng ra;
- **Esc** hoặc bấm ra ngoài là đóng.

Vì ô đó **chỉ đọc** nên gõ phím không sửa được dữ liệu — nó chỉ để đọc và copy. Muốn sửa giá trị trước khi ghi vào Excel thì vẫn nháy đúp ô **Giá trị** ở bảng chính như cũ.

Ô **Nhật ký** và các danh sách trong hộp thoại *Tạo file thư viện mới* cũng có menu chuột phải *Sao chép / Chọn tất cả*.

## 3m. Cửa sổ chạy mượt & icon trên Taskbar  *(1.7.5)*

### Kéo thả đổi kích thước không còn giật

Kéo vành cửa sổ trước đây khựng từng nấc. Không phải máy yếu — mà là mỗi pixel kéo
đi, Tk bắn ra một sự kiện `<Configure>`, và tool làm quá nhiều việc cho mỗi sự kiện
đó. Bản 1.7.5 cắt bốn chỗ:

| Chỗ | Trước | Nay |
|---|---|---|
| Sự kiện `<Configure>` của cửa sổ chính | Tk gửi sự kiện **lên theo `bindtags`**, mà bindtag của mọi widget con đều chứa đường dẫn của cửa sổ — nên một lệnh `bind` trên cửa sổ bắt luôn hàng trăm sự kiện của từng nút, từng ô nhập | Cửa sổ có **bindtag riêng** (`perf.tag_bind_self`), chỉ còn sự kiện của chính nó |
| Xếp lại ô tra cứu | Xếp lại `grid` mỗi lần bề rộng đổi 40 px | Bố cục chỉ có **hai dạng** (rộng ≥ 1000 px / hẹp) — **đổi dạng mới xếp lại**. Xếp lại `grid` là thứ đắt nhất trong cả chu trình vẽ |
| Thanh tab (Canvas) | Xoá sạch và vẽ lại mọi lúc, đo lại bề ngang từng tên tab bằng `font.measure()` | Gom sự kiện (40 ms), **nhớ kết quả đo**, và **bỏ qua hẳn** khi "chữ ký" lần vẽ (bề rộng, tab đang chọn, tab đang rê chuột, bảng màu) không đổi |
| Tab đã mở nhưng đang bị che | Vẫn nằm trong `grid`, Tk **tính lại bố cục cho cả 12 cây widget** mỗi lần cửa sổ đổi kích thước — mở càng nhiều tab, kéo càng nặng | `grid_remove()` khi chuyển tab; khung cuộn của tab bị ẩn cũng **ngừng tự đo** (`winfo_ismapped()`) |

Thêm hai chỗ nữa không liên quan tới kéo thả nhưng cũng nhanh lên rõ:

- **`autofit_tree` (dãn rộng cột bảng).** Bản cũ quét theo **cột** ở vòng ngoài,
  theo **dòng** ở vòng trong, nên `tree.item()` và `font.measure()` bị gọi
  *(số cột × số dòng)* lần — bảng 5.000 dòng 8 cột là **80.000 lệnh gọi** xuống
  Tcl, chính là vài giây chờ sau khi bấm *Tra cứu thư viện*. Nay quét **đúng một
  lượt**, mỗi cột giữ lại 3 chuỗi dài nhất (so sánh độ dài ký tự — phép tính
  thuần Python, gần như miễn phí) rồi mới đo pixel cho riêng chúng. Bề rộng cột
  ra y hệt.
- **Ô Nhật ký.** Trước đây mỗi dòng ghi ra là một lệnh `update_idletasks()` — vẽ
  lại **toàn bộ** cửa sổ. Ghi 500 dòng là 500 lần vẽ lại. Nay giới hạn nhiều nhất
  ~12 lần/giây; mắt không phân biệt được, nhưng các tác vụ dài nhanh lên thấy rõ.

Toàn bộ tiện ích gom / lọc sự kiện nằm ở **`perf.py`** — dùng lại được cho mọi
widget: `perf.on_resize()`, `perf.Debounce`, `perf.tag_bind_self()`, `perf.Throttle`.

### Icon trên Taskbar là ếch / nhện, không phải Python

Trước 1.7.5, cửa sổ đã đeo icon ếch trên thanh tiêu đề nhưng **nút dưới thanh
Taskbar vẫn là icon Python**. Lý do: Windows **không** lấy icon Taskbar từ cửa sổ.
Nó gom cửa sổ vào nút Taskbar theo **AppUserModelID** của tiến trình, rồi lấy icon
của ứng dụng đăng ký với ID đó. Chạy bằng `pythonw.exe` thì ID mặc định là của
Python — nên Windows vẽ icon Python, bất kể cửa sổ có icon gì.

Bản 1.7.5 làm hai việc:

1. `SetCurrentProcessExplicitAppUserModelID("Nhatnpm.AltiumDbTool.LibraryHelper")`
   ngay lúc nạp module, **trước khi cửa sổ đầu tiên được tạo**. Windows hết "mượn"
   Python vào và quay về dùng đúng icon của chính cửa sổ.
2. Gắn icon **thẳng vào HWND** bằng `WM_SETICON` — nạp riêng hai cỡ (`ICON_SMALL`
   cho thanh tiêu đề, `ICON_BIG` cho Taskbar và Alt+Tab) nên cả hai chỗ đều nét,
   không còn cảnh phóng to icon 16 px lên 32 px thành vỡ. Đặt lại một lần nữa ở
   sự kiện `<Map>` (lúc Windows báo cửa sổ đã hiện thật) cho chắc.

Đổi logo trong **Functions → Giao diện → Logo** vẫn ăn ngay lập tức ở cả ba chỗ:
ảnh lớn tab Liên hệ, icon thanh tiêu đề, và **nút Taskbar**.

> Chạy từ file `.exe` (mục **1b**) thì icon của **chính file exe** là icon đã nhúng
> lúc build — `AltiumDbTool-Frog.exe` ra ếch, `AltiumDbTool-Spider.exe` ra nhện —
> còn icon cửa sổ và Taskbar lúc chạy vẫn theo lựa chọn trong `config.json`.

## 3n. Cập nhật bản đóng gói `.exe`  *(1.7.5)*

Bản `.exe` **không** cập nhật được theo cách cũ. Hai lý do độc lập nhau, và cả hai
đều đã sửa ở 1.7.5:

**1. Updater từ chối gói không có `.py`.** `zip_info()` trước đây bắt buộc tìm thấy
`altium_db_tool.py` trong zip để đọc `APP_VERSION`. Zip chỉ chứa exe → báo *"không
phải bản cập nhật của tool"*. Nay updater nhận **hai kiểu gói**:

| Kiểu | Trong zip có | Số phiên bản đọc từ |
|---|---|---|
| `py` | `altium_db_tool.py` + các file mã nguồn | dòng `APP_VERSION` |
| `exe` | `AltiumDbTool.exe` + `version.json` | `version.json` |
| `exe+py` | cả hai | cả hai, **lệch nhau là báo lỗi và dừng** |

Không đọc được số phiên bản từ chính file exe, vì mã nguồn trong đó đã thành `.pyc`
— bên trong không còn dòng chữ `APP_VERSION = "1.7.5"` nào để tìm. Đó là lý do phải
có `version.json`.

Updater cũng **chặn cài chéo kiểu**: đang chạy exe mà gói chỉ có mã nguồn thì chép
`.py` vào cạnh exe cũng vô nghĩa (exe chạy code nằm bên trong chính nó); đang chạy
Python mà gói chỉ có exe thì cũng chẳng nâng cấp được gì. Cả hai chiều đều báo rõ
lý do thay vì cài xong rồi không thấy đổi gì.

**2. Windows khoá file exe đang chạy.** Đây mới là phần cốt lõi. `shutil.copy2()` đè
lên chính exe đang chạy ném `[WinError 32] The process cannot access the file
because it is being used by another process`. Lỗi này trước đây bị nuốt vào danh
sách `skipped`, nên tool báo *"đã cập nhật"* mà thực tế **thay 0 file** — đúng kiểu
lỗi âm thầm đã sửa ở 1.6.8, nay tái xuất ở dạng khác.

Windows cấm **ghi đè** exe đang chạy nhưng **vẫn cho đổi tên / di chuyển** nó — khoá
là khoá nội dung file, không khoá mục thư mục. Nên quy trình là:

```
1. đổi tên exe đang chạy  ->  _update_backup\<ngày_giờ>\AltiumDbTool.exe
2. chép exe mới vào đúng đường dẫn cũ
3. hỏi rồi tự khởi động lại
```

Hỏng ở bước 2 thì trả lại tên cũ ngay, máy không sứt mẻ gì. Chrome và VS Code cập
nhật y hệt cách này. Tiến trình đang chạy không hề gì khi file bị đổi tên: code đã
nằm trong bộ nhớ, và bản onefile của PyInstaller thì đã giải nén ra `%TEMP%` từ lúc
mở.

**Thứ tự có chủ ý:** file `.exe` được thay **trước tiên**, trước mọi file khác. Nếu
để nó chạy lẫn thì có cảnh này: thay exe thất bại nhưng `version.json` mới **vẫn**
được ghi đè, thành ra trên đĩa `version.json` nói "1.7.6" trong khi exe vẫn là
"1.7.5" — và bước kiểm lại sau khi cài đọc `version.json` nên báo *"cài thành
công"*. Một bản cập nhật thất bại mà báo thành công là kiểu lỗi tệ nhất.

**Sau khi cài xong**, tool hỏi *"Mở lại tool ngay?"* — Đồng ý thì nó chạy file exe
mới rồi thoát tiến trình cũ. Bắt buộc phải mở lại: exe trên đĩa đã là bản mới nhưng
tiến trình đang chạy vẫn là code cũ đã nạp vào bộ nhớ từ lúc mở.

**Hoàn tác** vẫn dùng được: exe cũ nằm nguyên trong `_update_backup\<ngày_giờ>\`,
bấm *"Hoàn tác bản cập nhật cuối"* là nó đổi chỗ ngược lại (cũng bằng phép đổi tên).
Bản mới lỡ hỏng không mở lên được thì vào thẳng thư mục đó, chép file exe cũ đè
ra ngoài bằng tay cũng xong.

**Những gì KHÔNG bị đụng tới** khi cập nhật, y như trước: `config.json`,
`rules.json`, `api_keys.enc`, `lookup_cache.db`, `state_log.db`. API key và cấu
hình giữ nguyên.

## 4. Tool xử lý file của bạn thế nào

Tool **học từ chính file Excel**, không hard-code:

| Loại cột | Cách xử lý |
|---|---|
| Cột có công thức (`="R"&B2`, `=A2`, `=Table10[[#This Row],...]`) | **Không ghi đè.** Copy công thức xuống dòng mới, để Excel tự tính — đúng như bạn đang làm tay. *1.7.3:* chỉ tính là cột công thức khi công thức có ở ≥ 60% số dòng và không trỏ sang file Excel khác — xem mục **3k** |
| Cột cố định (`Library Path`, `Footprint Path`, `Library Ref`...) | Lấy giá trị mà mọi dòng cũ đang dùng |
| `Package` / `Size` → `Footprint Ref` | Học từ các dòng cũ (0402→R0402, 0603→FB0603, 1206→F1206...). Sheet gộp nhiều loại thì khớp theo cặp `Type` + `Size`, không chắc thì để trống *(1.7.3)* |
| Thông số kỹ thuật | Lấy từ API và chuẩn hoá về đúng style của bạn: `1.5 kOhms`→`1.5K`, `0.1 µF`→`0.1UF`, `±1%`→`1%`, `0.063W, 1/16W`→`1/16W`, `0402 (1005 Metric)`→`0402` |
| `Design Item ID` / `Mfr. #`, `Manufacturer` / `Manufacturer 1`, `Description`, `Detailed Description`, `ComponentLink1URL` | Lấy thẳng từ API |
| `Type`, `Footprint type`, `Value` *(cột của kiểu Database_Library)* | Suy ra từ dữ liệu API: loại linh kiện từ Category/Description, `SMD`/`TH` từ `Mounting Type`, `Value` ghép thông số chính với định mức (`4.7K/1/8W`) *(1.7.3)* |

### Backup

Trước mỗi lần ghi, tool copy file Excel vào thư mục `_backup` cạnh file gốc, tên dạng `20260830_143512_Master_Library.xlsx`.

Nút **Backup...** trên thanh công cụ mở cửa sổ quản lý:

- danh sách các bản backup kèm thời điểm và dung lượng, tổng dung lượng ở dưới
- **Xoá mục đã chọn** (giữ Ctrl hoặc Shift để chọn nhiều dòng)
- **Xoá tất cả**
- **Chỉ giữ N bản mới nhất** — dọn nhanh, giữ lại vài bản gần đây
- **Mở thư mục** để xem bằng File Explorer

Mọi thao tác xoá đều hỏi xác nhận và cho biết sẽ giải phóng bao nhiêu dung lượng. Tool chỉ đụng vào các file backup của đúng file Excel đang mở — file khác trong thư mục đó không bị ảnh hưởng.

**Từ 1.6 backup dùng được cả hai chiều** (1.6.1: cửa sổ so sánh giờ đóng độc lập, không phải tắt cửa sổ quản lý backup trước):

- **Khôi phục bản đang chọn** — ghi đè file hiện tại bằng bản backup. Trước khi ghi đè, tool tự cất thêm một bản của file hiện tại tên `<ngày_giờ>_truoc-khi-khoi-phuc_Master_Library.xlsx`, nên khôi phục nhầm vẫn quay lại được. Nhớ đóng Excel trước.
- **So sánh với file hiện tại** — mở bảng khác biệt: dòng nào **thêm**, dòng nào **mất**, ô nào **sửa** (kèm giá trị cũ / mới). Các cột do công thức Excel tính được bỏ qua, nên không bị nhiễu bởi việc ghi qua openpyxl. Xuất CSV được.

Trong **Cấu hình API** có thêm ô *"Tự xoá backup cũ, chỉ giữ lại N bản mới nhất"* (mặc định 20, đặt 0 để giữ hết) — chạy tự động sau mỗi lần ghi.

### Dọn History của Altium

Mỗi lần bạn lưu một file `.SchLib` / `.PcbLib`, Altium nén bản cũ thành `History\<Tên>.~(n).SchLib.Zip`. Vài tháng là thư mục `History` to hơn cả thư viện thật.

Nút **History...** trên thanh công cụ mở cửa sổ dọn, dùng y hệt cửa sổ backup:

- tự tìm **mọi thư mục tên `History`** bên dưới gốc thư viện (`Symbols\History`, `Footprints\History`...)
- liệt kê từng file kèm **thư viện gốc**, thư mục, thời điểm, dung lượng
- **Chỉ giữ N bản mới nhất mỗi thư viện** — đếm riêng cho từng file `.SchLib`/`.PcbLib`, không phải chung cả thư mục
- các file sắp bị xoá tô **cam** ngay trong bảng để xem trước
- **Xoá mục đã chọn** (Ctrl/Shift chọn nhiều dòng) hoặc **Xoá tất cả**
- **Mở thư mục** bằng File Explorer

Tool **chỉ xoá file `.Zip` trong thư mục History**, không bao giờ đụng vào `.SchLib` / `.PcbLib` đang dùng. Mọi thao tác đều hỏi xác nhận và cho biết giải phóng bao nhiêu dung lượng. Số bản giữ mặc định đặt trong **Cấu hình API** (`max_history`, mặc định 10).

### Xuất CSV

Nút **Xuất CSV** ở các tab Tra cứu thư viện, Kiểm tra thư viện, Symbol/Footprint, Trùng thông số, Đối chiếu BOM không hỏi chỗ lưu nữa. File tự vào:

```
<Gốc xuất>\<tên báo cáo>\<tên báo cáo>_YYYYmmdd_HHMMSS.csv
```

Ví dụ `...\Exports\symbol_footprint\symbol_footprint_20260901_143512.csv`. Mỗi loại báo cáo một thư mục riêng, mỗi lần xuất một file có dấu ngày giờ nên không đè lên bản cũ, tiện so sánh theo thời gian. Xuất xong tool hỏi *"Mở thư mục chứa file?"*.

**Gốc xuất** đặt trong Cấu hình API (ô *Xuất CSV*); để trống thì dùng `<thư mục file Excel>\Exports`.

### Tải datasheet

Đặt **Thư mục lưu datasheet** trong Cấu hình API. Sau đó:

- Cửa sổ tra cứu → chọn một mã → **Tải datasheet** (không có link trong Excel thì tool tra online)
- Menu **Thư viện → Tải datasheet cho cả thư viện** — tải hàng loạt, bỏ qua file đã có

File lưu theo tên `<part number>.pdf`. Tool không tự sửa `ComponentLink1URL` trong Excel.

### Hai cách ghi file

- **Có `xlwings` + Excel** (mặc định trên Windows): tool điều khiển Excel thật → công thức được tính lại ngay, Table tự mở rộng, không mất gì.
- **Không có Excel**: dùng `openpyxl`. Công thức vẫn được copy xuống đúng, nhưng **giá trị đã lưu sẵn của các ô công thức cũ bị xoá** → mở file bằng Excel và bấm **Ctrl+S** một lần để Excel tính lại, rồi Altium mới đọc đủ.

## 5. Tuỳ biến — `rules.json`

File `rules.json` sinh ra lần chạy đầu, sửa được bằng Notepad:

```json
"Resistor": {
  "description": "RES {Resistance} OHM {Tolerance} {Power}  {Package}",
  "comment": "{Resistance}",
  "value_sort": { "from": "Resistance", "unit": "ohm" },
  "keywords": ["chip resistor", "resistor", "thick film"]
}
```

- `description` / `comment`: mẫu tạo nội dung — chỉ dùng cho cột **không có công thức**. `{Tên cột}` được thay bằng giá trị cột đó. Bộ lọc: `{Impedance @ Frequency|imp_short}` → `220R`.
- `value_sort`: cột nguồn + đơn vị (`ohm` hoặc `pf`).
- `keywords`: từ khoá để tự nhận diện loại linh kiện từ Category của Digi-Key. Thiếu loại nào cứ thêm.

Thêm sheet mới vào Excel → tool tự nhận và thêm mục tương ứng vào `rules.json`.

## 6. Kiểm tra nhanh

```
python self_test.py "D:\Altium\Library\Master_Library.xlsx"
python self_test.py "D:\Altium\Library\Master_Library.xlsx" --lint
```

Không tham số: mỗi sheet có cột nào là công thức, cột nào cố định, map Package→Footprint Ref, và thử điền 3 linh kiện mẫu — không cần API key.

`--lint` (1.6): chạy đúng bộ kiểm tra của tab *Sức khoẻ thư viện* ở chế độ dòng lệnh và trả mã thoát khác 0 nếu có **Lỗi** — dùng được trong hook trước khi commit hoặc trong GitHub Actions.

## 7. Lỗi thường gặp

| Lỗi | Xử lý |
|---|---|
| `Digi-Key OAuth lỗi 401` | Sai Client ID/Secret, hoặc app đang ở chế độ Sandbox — bỏ tick Sandbox |
| `khong tim thay '<pn>'` | Part number sai, hoặc Digi-Key không bán mã đó → thử nguồn Mouser |
| `File Excel đang mở trong Excel` | Đóng file rồi ghi lại (hoặc bật chế độ xlwings để ghi trực tiếp) |
| Cột trống sau khi ghi | Nhà cung cấp không có thông số đó — sửa tay trong bảng trước khi bấm Ghi |
| Altium đọc thấy ô trống | Mở Excel, Ctrl+S một lần (xem mục 4) |
| Panel Components trống sau khi clone repo | Tab **Sức khoẻ thư viện** → khung .DbLib → **Sửa đường dẫn theo máy này** |
| `Sai mật khẩu (hoặc file đã bị sửa đổi)` | Nhập sai mật khẩu file `api_keys.enc`. Quên hẳn thì xoá file đó rồi nhập lại key |
| `Không giải mã được bằng DPAPI` | File key được tạo trên tài khoản / máy khác. Dùng chế độ **Mật khẩu** nếu cần mang đi |
| Tra cứu ra số tồn kho cũ | Cache còn hạn — tick **Bỏ qua cache** ở tab Quét State, hoặc **Xoá cache** trong Cấu hình |

## 8. Cấu trúc thư mục

```
altium_db_tool.py   giao diện chính
providers.py        Digi-Key API v4 + Mouser Search API v2 (kèm giá theo bậc)
excel_db.py         đọc cấu trúc sheet, học quy luật, ghép dữ liệu
excel_writer.py     ghi Excel (xlwings hoặc openpyxl), backup / khôi phục / so sánh
db_format.py        1.7.3 - nhận diện kiểu file (Master_Library / Database_Library),
                    phân giải cột khoá và tên cột tương đương
                    1.7.4 - Master_Library là kiểu mặc định, Database_Library là tuỳ chọn
db_convert.py       1.7.3 - chuyển đổi qua lại giữa hai kiểu file
ui_dbconvert.py     1.7.3 - trang "Chuyển đổi định dạng DB" trong cửa sổ Functions
ui_copy.py          1.7.4 - bôi đen / Ctrl+C / menu chuột phải cho mọi bảng
normalize.py        chuẩn hoá đơn vị/ký hiệu
rules_default.py    quy luật mặc định cho từng sheet + cấu hình mặc định
AltiumDbTool.pyw    chạy không hiện cửa sổ đen
run_silent.vbs      chạy hoàn toàn im lặng
ui_tabs.py          thanh tab kiểu Chrome
ui_tools.py         các tab công cụ của 1.3-1.5
ui_tools2.py        các tab mới của 1.6 (Sức khoẻ, Giá BOM, Where-used, Thống kê)
lib_audit.py        quét State, kiểm tra symbol/footprint, trùng thông số
lib_lint.py         1.6 - soi lỗi dữ liệu Excel, symbol mồ côi, 3D model
altium_ole.py       trình đọc .SchLib/.PcbLib (OLE) tự viết, không cần olefile
altium_history.py   dọn thư mục History của Altium
bom_check.py        đọc và đối chiếu BOM
pricing.py          1.6 - giá theo bậc số lượng, MOQ, tính giá BOM
where_used.py       1.6 - đọc .SchDoc, biết mã nào dùng ở board nào
dblib_fix.py        1.6 - đọc / sửa .DbLib, kiểm tra ACE OLEDB
secure_store.py     1.6 - mã hoá API key (DPAPI hoặc mật khẩu)
cache.py            1.6 - cache tra cứu bằng SQLite
state_log.py        1.6 - nhật ký thay đổi State
stats.py            1.6 - thống kê + báo cáo HTML
updater.py          1.6.8 - kiểm tra bản mới ở GitHub Releases + cài từ file .zip
                    1.7.5 - cập nhật được cả bản .exe (đổi tên rồi thay), tự mở lại
theme.py            1.6 - bảng màu sáng / tối
perf.py             1.7.5 - gom / lọc sự kiện <Configure>, chống giật khi kéo
                    thả đổi kích thước cửa sổ (Debounce, on_resize,
                    tag_bind_self, Throttle)
template.py         1.6.3 - tạo file Excel + .DbLib mới đúng cấu trúc
ui_functions.py     1.6.7 - cửa sổ Functions + panel Backup/History/Cấu hình/Tab
ui_newlib.py        1.6.6 - hộp thoại tạo file thư viện mới
ui_widgets.py       1.6.7 - ô nhập gợi ý mờ, cột điều hướng, khung cuộn, hộp cập nhật
datasheets.py       tải datasheet
icon.ico/.png       logo con nhện (lựa chọn "Nhện")
frog.ico/.png       1.6.9 - logo con ếch (lựa chọn "Ếch", mặc định)
                    1.7.0 - hai file .ico ghi lại ở dạng BMP cho Tk đọc được
frog_icon.png       1.7.1 - bản cắt sát, dùng cho icon cửa sổ cỡ nhỏ
self_test.py        kiểm tra không cần API
build_exe.bat       1.7.5 - double-click để build exe + đóng gói bản phát hành
AltiumDbTool.spec   1.7.5 - cấu hình PyInstaller, sinh AltiumDbTool.exe (icon ếch)
make_release.py     1.7.5 - sinh version.json + History/AltiumDbTool_v*_exe.zip
version_info.txt    1.7.5 - tên, phiên bản, tác giả hiện trong Properties của exe
config.json         đường dẫn + tuỳ chọn (tự sinh, KHÔNG còn chứa API key)
api_keys.enc        1.6 - API key đã mã hoá (tự sinh)
rules.json          quy luật, sửa được (tự sinh)
version.json        1.7.5 - số phiên bản của bản đóng gói, nằm cạnh file exe
lookup_cache.db     1.6 - cache tra cứu (tự sinh, xoá được)
state_log.db        1.6 - nhật ký State (tự sinh)
```

---


