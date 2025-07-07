var DocTienBangChu = function () {
    this.ChuSo = new Array(" không ", " một ", " hai ", " ba ", " bốn ", " năm ", " sáu ", " bảy ", " tám ", " chín ");
    this.Tien = new Array("", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ");
};

DocTienBangChu.prototype.docSo3ChuSo = function (baso) {
    var tram;
    var chuc;
    var donvi;
    var KetQua = "";
    tram = parseInt(baso / 100);
    chuc = parseInt((baso % 100) / 10);
    donvi = baso % 10;
    if (tram == 0 && chuc == 0 && donvi == 0) return "";
    if (tram != 0) {
        KetQua += this.ChuSo[tram] + " trăm ";
        if ((chuc == 0) && (donvi != 0)) KetQua += " linh ";
    }
    if ((chuc != 0) && (chuc != 1)) {
        KetQua += this.ChuSo[chuc] + " mươi";
        if ((chuc == 0) && (donvi != 0)) KetQua = KetQua + " linh ";
    }
    if (chuc == 1) KetQua += " mười ";
    switch (donvi) {
        case 1:
            if ((chuc != 0) && (chuc != 1)) {
                KetQua += " mốt ";
            }
            else {
                KetQua += this.ChuSo[donvi];
            }
            break;
        case 5:
            if (chuc == 0) {
                KetQua += this.ChuSo[donvi];
            }
            else {
                KetQua += " lăm ";
            }
            break;
        default:
            if (donvi != 0) {
                KetQua += this.ChuSo[donvi];
            }
            break;
    }
    return KetQua;
}

DocTienBangChu.prototype.doc = function (SoTien) {
    var lan = 0;
    var i = 0;
    var so = 0;
    var KetQua = "";
    var tmp = "";
    var soAm = false;
    var ViTri = new Array();
    if (SoTien < 0) soAm = true;//return "Số tiền âm !";
    if (SoTien == 0) return "Không đồng";//"Không đồng !";
    if (SoTien > 0) {
        so = SoTien;
    }
    else {
        so = -SoTien;
    }
    if (SoTien > 8999999999999999) {
        //SoTien = 0;
        return "";//"Số quá lớn!";
    }
    ViTri[5] = Math.floor(so / 1000000000000000);
    if (isNaN(ViTri[5]))
        ViTri[5] = "0";
    so = so - parseFloat(ViTri[5].toString()) * 1000000000000000;
    ViTri[4] = Math.floor(so / 1000000000000);
    if (isNaN(ViTri[4]))
        ViTri[4] = "0";
    so = so - parseFloat(ViTri[4].toString()) * 1000000000000;
    ViTri[3] = Math.floor(so / 1000000000);
    if (isNaN(ViTri[3]))
        ViTri[3] = "0";
    so = so - parseFloat(ViTri[3].toString()) * 1000000000;
    ViTri[2] = parseInt(so / 1000000);
    if (isNaN(ViTri[2]))
        ViTri[2] = "0";
    ViTri[1] = parseInt((so % 1000000) / 1000);
    if (isNaN(ViTri[1]))
        ViTri[1] = "0";
    ViTri[0] = parseInt(so % 1000);
    if (isNaN(ViTri[0]))
        ViTri[0] = "0";
    if (ViTri[5] > 0) {
        lan = 5;
    }
    else if (ViTri[4] > 0) {
        lan = 4;
    }
    else if (ViTri[3] > 0) {
        lan = 3;
    }
    else if (ViTri[2] > 0) {
        lan = 2;
    }
    else if (ViTri[1] > 0) {
        lan = 1;
    }
    else {
        lan = 0;
    }
    for (i = lan; i >= 0; i--) {
        tmp = this.docSo3ChuSo(ViTri[i]);
        KetQua += tmp;
        if (ViTri[i] > 0) KetQua += this.Tien[i];
        if ((i > 0) && (tmp.length > 0)) KetQua += '';//',';//&& (!string.IsNullOrEmpty(tmp))
    }
    if (KetQua.substring(KetQua.length - 1) == ',') {
        KetQua = KetQua.substring(0, KetQua.length - 1);
    }
    KetQua = KetQua.substring(1, 2).toUpperCase() + KetQua.substring(2);
    if (soAm) {
        return "Âm " + KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
    else {
        return KetQua + " đồng";//.substring(0, 1);//.toUpperCase();// + KetQua.substring(1);
    }
}

function formatNumber(numberString) {
    if (!numberString) return '';
    // Loại bỏ tất cả dấu chấm
    const num = numberString.replace(/\./g, '');
    const formatted = parseFloat(num).toString();
    return formatted.replace('.', ',');
}

const SPREADSHEET_ID_1 = '1_VjjzKwaUdjxsOPdLOGHzcc0b8oyQsj4_Duw_7xmmWo';
const RANGE_1 = 'tong_hop_don_hang_24!A:AU';
const RANGE_CHITIET_1 = 'tong_hop_don_hang_chi_tiet_24!C:Y';


const SPREADSHEET_ID_2 = '1FLsjyTBi_JfkcDomkgq1ChLWUtrnKv-PfGpPJwt7Zek';
const RANGE_2 = 'tong_hop_don_hang_25!A:AU';
const RANGE_CHITIET_2 = 'tong_hop_don_hang_chi_tiet_25!C:Y';

const API_KEY = 'AIzaSyA9g2qFUolpsu3_HVHOebdZb0NXnQgXlFM';

// Lấy giá trị từ URI sau dấu "?" cho các tham số cụ thể
function getDataFromURI() {
    const url = window.location.href;

    // Sử dụng RegEx để trích xuất giá trị của ma_don_hang và QRCODE
    const maDonHangURIMatch = url.match(/ma_don_hang=([^?&]*)/);
    const maKhachHangURIMatch = url.match(/ma_khach_hang=([^?&]*)/);
    const qrCodeMatch = url.match(/QRCODE=(.*)$/);  // Sử dụng regex để lấy tất cả sau QRCODE=

    // Gán các giá trị vào các biến
    const maDonHangURI = maDonHangURIMatch ? decodeURIComponent(maDonHangURIMatch[1]) : null;
    const maKhachHangURI = maKhachHangURIMatch ? decodeURIComponent(maKhachHangURIMatch[1]) : null;
    const qrCode = qrCodeMatch ? decodeURIComponent(qrCodeMatch[1]) : null;

    // Trả về một đối tượng chứa các giá trị
    return {
        maDonHangURI,
        maKhachHangURI,
        qrCode
    };
}

// Hàm để tải Google API Client
function loadGapiAndInitialize() {
    const script = document.createElement('script');
    script.src = "https://apis.google.com/js/api.js"; // Đường dẫn đến Google API Client
    script.onload = initialize; // Gọi hàm `initialize` sau khi thư viện được tải xong
    script.onerror = () => console.error('Failed to load Google API Client.');
    document.body.appendChild(script); // Gắn thẻ script vào tài liệu
}

// Hàm khởi tạo sau khi Google API Client được tải
function initialize() {
    gapi.load('client', async () => {
        try {
            await gapi.client.init({
                apiKey: API_KEY,
                discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4']
            });

            const uriData = getDataFromURI();
            if (!uriData.maDonHangURI || !uriData.qrCode) {
                updateContent('No valid data found in URI.');
                return;
            }

            findRowInSheet(uriData.maDonHangURI);
            findDetailsInSheet(uriData.maDonHangURI);

            // Cập nhật nội dung hoặc xử lý thêm thông tin QR Code
            updateQRCodeContent(uriData.qrCode);

        } catch (error) {
            updateContent('Initialization error: ' + error.message);
            console.error('Initialization Error:', error);
        }
    });
}

// Gọi hàm tải Google API Client khi DOM đã sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    loadGapiAndInitialize();
});

function updateQRCodeContent(qrCode) {
    // Gắn QR Code vào trong nội dung trang (VD: hiển thị ảnh QR code)
    const qrCodeElement = document.getElementById('qr-code');
    if (qrCodeElement) {
        qrCodeElement.src = qrCode;
        // Đặt kích thước của QR Code
        qrCodeElement.style.width = '300px';  // Chiều rộng 150px
        qrCodeElement.style.height = 'auto';  // Chiều cao tự động
    }
}

function updateContent(message) {
    const contentElement = document.getElementById('content'); // Thay 'content' bằng ID của phần tử HTML cần hiển thị
    if (contentElement) {
        contentElement.textContent = message;
    } else {
        console.warn('Element with ID "content" not found.');
    }
}


// Tìm chỉ số dòng chứa dữ liệu khớp trong cột B và lấy các giá trị từ các cột khác
let orderDetails = null; // Thông tin đơn hàng chính
let orderItems = [];

async function findRowInSheet(maDonhangURI) {
    try {
        // Tìm trong bảng tính đầu tiên
        const found = await searchInSheet(SPREADSHEET_ID_1, RANGE_1, maDonhangURI);
        if (found) return;

        // Nếu không tìm thấy, tìm trong bảng tính thứ hai
        const foundInSecondSheet = await searchInSheet(SPREADSHEET_ID_2, RANGE_2, maDonhangURI);
        if (!foundInSecondSheet) {
            updateContent(`No matching data found for "${maDonhangURI}" in both spreadsheets.`);
        }
    } catch (error) {
        updateContent('Error fetching data: ' + error.message);
        console.error('Fetch Error:', error);
    }
}

async function searchInSheet(spreadsheetId, range, maDonhangURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: range,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            return false;
        }

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];

            const bColumnValue = row[1]; // Cột B
            const uriData = getDataFromURI();
            if (bColumnValue === maDonhangURI) {
                // Lưu dữ liệu vào biến toàn cục
                orderDetails = {
                    phuongThucban: row[0] || '', // Cột A
                    maDonhang: row[1] || '', // Cột B
                    maHopdong: row[46] || '', // Cột AU
                    donviPhutrach: row[5] || '', // Cột F
                    tenNguoilienhe: row[13] || '', // Cột N
                    tenKhachhangcuoi: row[23] || '', // Cột X
                    tenTochuc: row[15] || '', // Cột P
                    diachi: row[8] || '', // Cột I
                    diachiChitiet: row[17] || '', // Cột R
                    diachiKhachhangcuoi: row[24] || '', // Cột Y
                    tenNhanvien: row[4] || '', // Cột E
                    sdtNhanvien: row[6] || '', // Cột G
                    sdtKhachhang: row[18] || '', // Cột S
                    sdtKhachhangcuoi: row[22] || '', // Cột W
                    emailKhachhang: row[19] || '', // Cột T
                    hanGiaohang: row[50] || '', // Cột AY
                    tongSobo: row[25] || '', // Cột Z
                    congnpp: row[35] || '', // Cột AJ
                    mucChietkhaunpp: row[36] || '', // Cột AK
                    giatriChietkhaunpp: row[37] || '', // Cột AL
                    phiVanchuyenlapdatnpp: row[38] || '', // Cột AM
                    mucthueGTGTnpp: row[39] || '', // Cột AN
                    thueGTGTnpp: row[40] || '', // Cột AO
                    tamUngnpp: row[41] || '', // Cột AP
                    sotienPhaithanhtoannpp: row[42] || '', // Cột AQ
                    maKhachHang: uriData.maKhachHangURI,
                };

                // Xử lý dữ liệu tìm được
                processFoundData(orderDetails);
                return true; // Dừng khi tìm thấy
            }
        }
        return false; // Không tìm thấy
    } catch (error) {
        console.error('Error in searchInSheet:', error);
        return false;
    }
}

function processFoundData(orderDetails) {
    // Tính toán giá trị sotienConthieu
    const sotienConthieunpp = formatNumber(orderDetails.sotienPhaithanhtoannpp) - formatNumber(orderDetails.tamUngnpp);

    // Định dạng và chuyển đổi số tiền
    const doctien = new DocTienBangChu();
    const sotienBangchu = doctien.doc(sotienConthieunpp);

    // Cập nhật giá trị sotienBangchu vào orderDetails
    orderDetails.sotienConthieunpp = sotienConthieunpp.toLocaleString('vi-VN', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    orderDetails.sotienBangchu = sotienBangchu;

    // Cập nhật DOM
    Object.keys(orderDetails).forEach((key) => {
        if (orderDetails[key]) {
            updateElement(key, orderDetails[key]);
        }
    });
    if (sotienBangchu) updateElement('sotienBangchu', sotienBangchu);

    // Kiểm tra và hiển thị mã hợp đồng
    if (orderDetails.maHopdong && orderDetails.maHopdong.trim() !== '') {
        const hopdongContainer = document.getElementById('hopdong-container');
        document.getElementById('maHopdong').textContent = orderDetails.maHopdong;
        hopdongContainer.style.display = 'block';
    }

    // Gọi các hàm hiển thị nội dung
    displayHTML(orderDetails);
    displayConditions(orderDetails);

    function toggleRowVisibility(rowId, value) {
        const row = document.getElementById(rowId);
        if (row) {
            const stringValue = typeof value === 'string' ? value : String(value || '');
            const numericValue = parseFloat(stringValue.replace(/\./g, '').replace(',', '.') || '0');
            row.style.display = numericValue > 0 ? '' : 'none';
        }
    }

    // Ẩn/hiện các dòng theo điều kiện
    toggleRowVisibility('rowChietKhau', orderDetails.giatriChietkhaunpp);
    toggleRowVisibility('rowPhiVanChuyen', orderDetails.phiVanchuyenlapdatnpp);
    toggleRowVisibility('rowThueGTGT', orderDetails.thueGTGTnpp);
    toggleRowVisibility('rowTamUng', orderDetails.tamUngnpp);

    // Hiển thị hoặc ẩn nội dung thanh toán
    const paymentContent = document.getElementById('payment-content');
    if (paymentContent) {
        paymentContent.style.display =
            orderDetails.donviPhutrach === "BP. BH1" || orderDetails.donviPhutrach === "BP. BH2"
                ? 'block'
                : 'none';
    }
}


function displayHTML() {
    // Trích xuất các giá trị cần thiết từ data
    const phuongThucban = orderDetails.phuongThucban || '';
    const donviPhutrach = orderDetails.donviPhutrach || '';
    const tenNguoilienhe = orderDetails.tenNguoilienhe || '';
    const tenKhachhangcuoi = orderDetails.tenKhachhangcuoi || '';
    const tenTochuc = orderDetails.tenTochuc || '';
    const diachi = orderDetails.diachi || '';
    const diachiChitiet = orderDetails.diachiChitiet || '';
    const diachiKhachhangcuoi = orderDetails.diachiKhachhangcuoi || '';
    const tenNhanvien = orderDetails.tenNhanvien || '';
    const sdtNhanvien = orderDetails.sdtNhanvien || '';
    const sdtKhachhang = orderDetails.sdtKhachhang || '';
    const sdtKhachhangcuoi = orderDetails.sdtKhachhangcuoi || '';
    const emailKhachhang = orderDetails.emailKhachhang || '';
    const hanGiaohang = orderDetails.hanGiaohang || '';
    const today = new Date();
    const ngayPhatHanh = today.toLocaleDateString('vi-VN');
    if (ngayPhatHanh) updateElement('ngayPhatHanh', ngayPhatHanh);
    // Cập nhật giá trị ngayPhatHanh vào orderDetails
    orderDetails.ngayPhatHanh = ngayPhatHanh;


    let htmlContent = "";

    if (donviPhutrach === "BP. BH1" && phuongThucban !== "Bán chéo") {
        htmlContent = `
                            <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Kính gửi:</i></td>
                                        <td class="infocol-2">${tenNguoilienhe}</td>
                                        <td class="infocol-3"><i>Ngày phát hành:</i></td>
                                        <td class="infocol-4">${ngayPhatHanh}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Đơn vị:</i></td>
                                        <td class="infocol-2">${tenTochuc}</td>
                                        <td class="infocol-3"><i>Đơn vị trực thuộc:</i></td>
                                        <td class="infocol-4">${donviPhutrach}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Địa chỉ:</i></td>
                                        <td class="infocol-2">${diachiChitiet}</td>
                                        <td class="infocol-3"><i>Soạn báo giá:</i></td>
                                        <td class="infocol-4">${tenNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>SĐT:</i></td>
                                        <td class="infocol-2">${sdtKhachhang}</td>
                                        <td class="infocol-3"><i>SĐT:</i></td>
                                        <td class="infocol-4">${sdtNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Email:</i></td>
                                        <td class="infocol-2">${emailKhachhang}</td>
                                        <td class="infocol-3"><i>CSKH:</i></td>
                                        <td class="infocol-4">
                                                <b><font color="red">1900 0282</font></b>
                                            </td>
                                    </tr>
                                </tbody>
                            `;
    } else if (donviPhutrach === "BP. BH1" && phuongThucban === "Bán chéo") {
        htmlContent = `
                                <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Kính gửi:</i></td>
                                        <td class="infocol-2">${tenNguoilienhe}</td>
                                        <td class="infocol-3"><i>Ngày phát hành:</i></td>
                                        <td class="infocol-4">${ngayPhatHanh}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Đơn vị:</i></td>
                                        <td class="infocol-2"></td>
                                        <td class="infocol-3"><i>Đơn vị trực thuộc:</i></td>
                                        <td class="infocol-4">${donviPhutrach}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Địa chỉ:</i></td>
                                        <td class="infocol-2">${diachiChitiet}</td>
                                        <td class="infocol-3"><i>Soạn báo giá:</i></td>
                                        <td class="infocol-4">${tenNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>SĐT:</i></td>
                                        <td class="infocol-2">${sdtKhachhang}</td>
                                        <td class="infocol-3"><i>SĐT:</i></td>
                                        <td class="infocol-4">${sdtNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Địa chỉ công trình:</i></td>
                                        <td class="infocol-2">
                                                <b><font color="red">${tenKhachhangcuoi} - ${diachiKhachhangcuoi} - ${sdtKhachhangcuoi}</font></b>
                                            </td>
                                        <td class="infocol-3"><i>CSKH:</i></td>
                                        <td class="infocol-4">
                                                <b><font color="red">1900 0282</font></b>
                                            </td>
                                    </tr>
                                </tbody>
                                `;
    } else if (donviPhutrach !== "BP. BH1" && hanGiaohang !== "") {
        htmlContent = `
                                   <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Kính gửi:</i></td>
                                        <td class="infocol-2"></span></td>
                                        <td class="infocol-3"><i>Ngày phát hành:</i></td>
                                        <td class="infocol-4">${ngayPhatHanh}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Đơn vị:</i></td>
                                        <td class="infocol-2">${donviPhutrach}</td>
                                        <td class="infocol-3"><i>Dự kiến giao:</i></td>
                                        <td class="infocol-4">${hanGiaohang}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Địa chỉ:</i></td>
                                        <td class="infocol-2">${diachi}</td>
                                        <td class="infocol-3"><i>Soạn báo giá:</i></td>
                                        <td class="infocol-4">${tenNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>SĐT:</i></td>
                                        <td class="infocol-2">${sdtNhanvien}</td>
                                        <td class="infocol-3"><i></i></td>
                                        <td class="infocol-4"></td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Công trình:</i></td>
                                        <td class="infocol-2">
                                                <b><font color="red">${tenNguoilienhe} - ${diachiChitiet} - ${sdtKhachhang}</font></b>
                                            </td>
                                        <td class="infocol-3"><i>CSKH:</i></td>
                                        <td class="infocol-4">
                                                <b><font color="red">1900 0282</font></b>
                                        </td>
                                    </tr>
                                </tbody>
                                `;
    } else if (donviPhutrach !== "BP. BH1" && hanGiaohang === "") {
        htmlContent = `
                                   <tbody>
                                    <tr>
                                        <td class="infocol-1"><i>Kính gửi:</i></td>
                                        <td class="infocol-2"></span></td>
                                        <td class="infocol-3"><i>Ngày phát hành:</i></td>
                                        <td class="infocol-4">${ngayPhatHanh}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Đơn vị:</i></td>
                                        <td class="infocol-2">${donviPhutrach}</td>
                                        <td class="infocol-3"><i>Dự kiến giao:</i></td>
                                        <td class="infocol-4">Trao đổi với QLSX</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Địa chỉ:</i></td>
                                        <td class="infocol-2">${diachi}</td>
                                        <td class="infocol-3"><i>Soạn báo giá:</i></td>
                                        <td class="infocol-4">${tenNhanvien}</td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>SĐT:</i></td>
                                        <td class="infocol-2">${sdtNhanvien}</td>
                                        <td class="infocol-3"><i></i></td>
                                        <td class="infocol-4"></td>
                                    </tr>
                                    <tr>
                                        <td class="infocol-1"><i>Công trình:</i></td>
                                        <td class="infocol-2">
                                                <b><font color="red">${tenNguoilienhe} - ${diachiChitiet} - ${sdtKhachhang}</font></b>
                                            </td>
                                        <td class="infocol-3"><i>CSKH:</i></td>
                                        <td class="infocol-4">
                                                <b><font color="red">1900 0282</font></b>
                                        </td>
                                    </tr>
                                </tbody>
                                `;
    }

    document.getElementById("content").innerHTML = htmlContent;
}

function displayConditions() {
    // Trích xuất các giá trị cần thiết từ data
    const phuongThucban = orderDetails.phuongThucban || '';
    const donviPhutrach = orderDetails.donviPhutrach || '';
    const thueGTGTnpp = orderDetails.thueGTGTnpp || '';
    const phiVanchuyenlapdatnpp = orderDetails.phiVanchuyenlapdatnpp || '';

    let outputHTML = ""; // Đổi tên biến từ htmlContent thành outputHTML

    // Điều kiện cho thuế GTGT
    if (donviPhutrach !== "BP. BH1") {
        outputHTML += `<p>1. Giá trên đã bao gồm thuế GTGT.</p>`;
    } else {
        outputHTML += `<p>1. Giá trên đã bao gồm thuế GTGT.</p>`;
    }

    // Điều kiện cho phí vận chuyển và phương thức bán
    if ((phiVanchuyenlapdatnpp === "" || phiVanchuyenlapdatnpp === "0") && donviPhutrach === "BP. BH1" && phuongThucban == "Bán lẻ") {
        outputHTML += `<p>2. Giá trên đã bao gồm phí vận chuyển, lắp đặt tại nội thành Hà Nội.</p>`;
    } else if (phiVanchuyenlapdatnpp !== '' && phiVanchuyenlapdatnpp !== "0" && donviPhutrach !== "BP. BH1") {
        outputHTML += `<p>2. Giá trên đã bao gồm phí vận chuyển, lắp đặt.</p>`;
    } else {
        outputHTML += `<p>2. Giá trên chưa bao gồm phí vận chuyển, lắp đặt.</p>`;
    }

    // Điều kiện cho đơn vị phụ trách và phương thức bán
    if (donviPhutrach === "BP. BH1" && phuongThucban === "Bán đại lý") {
        outputHTML += `
                                <p>3. Giá trên có hiệu lực 30 ngày kể từ ngày phát hành.</p>
                                <p>4. Thanh toán 100% tổng giá trị đơn hàng trước khi giao hàng.</p>
                                <p>5. Giao hàng sau 3 đến 5 ngày làm việc, kể từ ngày chốt đơn.</p>
                                <p>6. Thời gian bảo hành 12 tháng theo tiêu chuẩn của nhà sản xuất.</p>
                            `;
    } else if (donviPhutrach !== "BP. BH1") {
        outputHTML += `
                                <p>3. Giá trên có hiệu lực 30 ngày kể từ ngày phát hành.</p>
                                <p>4. Tạm ứng 50% tổng giá trị đơn hàng, thanh toán hết số còn lại sau khi nghiệm thu bàn giao.</p>
                                <p>5. Giao hàng sau 3 đến 5 ngày làm việc.</p>
                                <p>6. Thời gian bảo hành:</p>
                                <p style="padding-left: 20px;"> - Bảo hành 2 năm phụ kiện lắp đồng bộ với cửa, tay nắm, bánh xe, ổ khóa, động cơ, điều khiển, phụ kiện nhựa, …</p>
                                <p style="padding-left: 20px;"> - Bảo hành 2 năm sản phẩm Cửa xếp rèm tổ ong; Cửa xếp 2 trong 1/ xếp xích 2 trong 1 kết hợp lưới và rèm.</p>
                                <p style="padding-left: 20px;"> - Bảo hành 3 năm sản phẩm Cửa lưới sợi thủy tinh; Cửa xếp/ xếp xích lưới PL, PVC, PET; Cửa xếp lưới nhôm, lá nhựa PC.</p>
                                <p style="padding-left: 20px;"> - Bảo hành 5 năm sản phẩm Cửa lưới Inox.</p>
                            `;
    } else {
        outputHTML += `
                                <p>3. Giá trên có hiệu lực 30 ngày kể từ ngày phát hành.</p>
                                <p>4. Tạm ứng 50% tổng giá trị đơn hàng, thanh toán hết số còn lại sau khi nghiệm thu bàn giao.</p>
                                <p>5. Lắp đặt sau 5 đến 7 ngày làm việc, kể từ ngày nhận được tiền tạm ứng lần 1.</p>
                                <p>6. Thời gian bảo hành:</p>
                                <p style="padding-left: 20px;"> - Bảo hành 2 năm phụ kiện lắp đồng bộ với cửa, tay nắm, bánh xe, ổ khóa, động cơ, điều khiển, phụ kiện nhựa, …</p>
                                <p style="padding-left: 20px;"> - Bảo hành 2 năm sản phẩm Cửa xếp rèm tổ ong; Cửa xếp 2 trong 1/ xếp xích 2 trong 1 kết hợp lưới và rèm.</p>
                                <p style="padding-left: 20px;"> - Bảo hành 3 năm sản phẩm Cửa lưới sợi thủy tinh; Cửa xếp/ xếp xích lưới PL, PVC, PET; Cửa xếp lưới nhôm, lá nhựa PC.</p>
                                <p style="padding-left: 20px;"> - Bảo hành 5 năm sản phẩm Cửa lưới Inox.</p>
                            `;
    }

    // Hiển thị nội dung HTML trong thẻ `content`
    document.getElementById("contentfooter").innerHTML = outputHTML;
}

// Tìm chi tiết trong bảng tính
async function findDetailsInSheet(maDonhangURI) {
    try {
        // Tìm trong bảng tính đầu tiên
        const found = await searchDetailsInSheet(SPREADSHEET_ID_1, RANGE_CHITIET_1, maDonhangURI);
        if (found) return;

        // Nếu không tìm thấy, tìm trong bảng tính thứ hai
        const foundInSecondSheet = await searchDetailsInSheet(SPREADSHEET_ID_2, RANGE_CHITIET_2, maDonhangURI);
        if (!foundInSecondSheet) {
            updateContent('No matching detail data found in both spreadsheets.');
        }
    } catch (error) {
        console.error('Error fetching detail data:', error);
        updateContent('Error fetching detail data.');
    }
}

async function searchDetailsInSheet(spreadsheetId, range, maDonhangURI) {
    try {
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: range,
        });

        const rows = response.result.values;
        if (!rows || rows.length === 0) {
            return false;
        }

        const filteredRows = rows.filter(row => row[0] === maDonhangURI); // Lọc các dòng có giá trị cột A khớp với maDonhangURI
        if (filteredRows.length > 0) {
            orderItems = filteredRows.map(extractDetailDataFromRow);
            displayDetailData(filteredRows);
            return true; // Dừng khi tìm thấy
        }
        return false; // Không tìm thấy
    } catch (error) {
        console.error('Error in searchDetailsInSheet:', error);
        return false;
    }
}

function displayDetailData(filteredRows) {
    const tableBody = document.getElementById('itemTableBody');
    tableBody.innerHTML = ''; // Xóa dữ liệu cũ nếu có

    filteredRows.forEach(row => {
        const item = extractDetailDataFromRow(row);;

        // Kiểm tra giá trị của donviPhutrach
        if (!item.maDonhangCT.includes("1C.029.01")) {
            tableBody.innerHTML += `
                <tr class="bordered-table">
                    <td class="borderedcol-1">${item.sttTrongdon || ''}</td>
                    <td class="borderedcol-2">${item.vitriLapdat || ''}</td>
                    <td class="borderedcol-3">${item.maSanphamid || ''}</td>
                    <td class="borderedcol-4">${item.ghiChu ? `${item.diengiai || ''} - ${item.ghiChu}` : item.diengiai || ''}</td>
                    <td class="borderedcol-5">${item.chieuRong || ''}</td>
                    <td class="borderedcol-6">${item.chieuCao || ''}</td>
                    <td class="borderedcol-7">${item.dienTich || ''}</td>
                    <td class="borderedcol-8">${item.soLuong || ''}</td>
                    <td class="borderedcol-9">${item.dvt || ''}</td>
                    <td class="borderedcol-10">${item.khoiLuong || ''}</td>
                    <td class="borderedcol-11">${item.dongianpp || ''}</td>
                    <td class="borderedcol-12">${item.giabannpp || ''}</td>
                </tr>
            `;
        } else {
            tableBody.innerHTML += `
                <tr class="bordered-table">
                    <td class="borderedcol-1">${item.sttTrongdon || ''}</td>
                    <td class="borderedcol-2">${item.vitriLapdat || ''}</td>
                    <td class="borderedcol-3">${item.maSanphamid || ''}</td>
                    <td class="borderedcol-4">${item.diengiai || ''}</td>
                    <td class="borderedcol-5">${item.chieuRong || ''}</td>
                    <td class="borderedcol-6">${item.chieuCao || ''}</td>
                    <td class="borderedcol-7">${item.dienTich || ''}</td>
                    <td class="borderedcol-8">${item.soLuong || ''}</td>
                    <td class="borderedcol-9">${item.dvt || ''}</td>
                    <td class="borderedcol-10">${item.khoiLuong || ''}</td>
                    <td class="borderedcol-11">${item.dongianpp || ''}</td>
                    <td class="borderedcol-12">${item.giabannpp || ''}</td>
                </tr>
            `;
        }
    });
}


// Trích xuất dữ liệu từ hàng
function extractDetailDataFromRow(row) {
    return {
        maDonhangCT: row[0],
        sttTrongdon: row[1],
        vitriLapdat: row[3],
        diengiai: row[10],
        maSanphamid: row[8],
        ghiChu: row[17],
        chieuRong: row[11],
        chieuCao: row[12],
        dienTich: row[13],
        soLuong: row[14],
        dvt: row[15],
        khoiLuong: row[16],
        dongianpp: row[20],
        giabannpp: row[21],
    };
}

// Hàm cập nhật nội dung DOM
function updateElement(elementId, value) {
    const element = document.getElementById(elementId);
    if (element) {
        element.innerText = value;
    }
}