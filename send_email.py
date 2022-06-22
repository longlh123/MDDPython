import smtplib

#A message as a string formatted as specified in the various RFCs
message = """From: VN-Survey <vn-survey@ipsosresearch.com>
To: Long.Pham <long.pham@ipsos.com>
MIME-Version: 1.0
Content-type: text/html
Subject: SMTP HTML e-mail test

<center>
    <table style='border-collapse:collapse; width:795px;'>
        <tr>
            <td style='padding-top:10px; width:100%; background-color: rgb(45, 150, 205);'><td>
        </tr>
        <tr style='height:450px'>
            <td  style='padding:20px 20px; width:100%; background-image: url(https://images1.ipsosinteractive.com/ABC_VIETNAM_10072020/images/ford_header.png); background-repeat: no-repeat; background-color: rgb(45, 150, 205); position: relative; '>
                <span style='position:absolute; top: 300px; color:white;font-size:14pt;font-family:Arial,sans-serif;'>
                    <h3>Xin chào Aaron McCune Stein,</h3>
                    <h2>Cảm ơn bạn đã chọn Ford Everest!</h2>
                </span>
            <td>
        </tr>
        <tr>
            <td  style='padding:30px 40px; width:100%; background-color: rgb(0, 52, 120); color:white;font-size:16pt;font-family:Arial,sans-serif;line-height:150%; text-align:center'>
                Chúng tôi sẽ đánh giá cao nếu bạn dành vài phút để cho chúng tôi biết về chiếc Ford Everest của bạn, được mua vào khoảng tháng Tháng Giêng 2022. Phản hồi của bạn là rất quan trọng đối với chúng tôi!
            <td>
        </tr>
        <tr style='background-color:#F2F2F2'>
            <td  style='padding:20px 20px; width:100%; color:#003478;font-size:14pt;font-family:Arial,sans-serif;'>
                <ul>
                    <li>Link sau sẽ dẫn quý khách đến phần khảo sát: <a href='https://ford-quality.com.vn/vg'>https://ford-quality.com.vn/vg</a></li>
                    <li>Sau đó, quý khách nhấn "Tiếp Tục"</li>
                    <li>Xin quý khách vui lòng hoàn thành bản khảo sát trước ngày 12/05/2022</li>
                </ul>
            <td>
        </tr>
        <tr style='background-color:#F2F2F2'>
            <td  style='padding:20px 20px; width:100%; font-weight:24px; color:#003478;font-size:14pt;font-family:Arial,sans-serif;line-height:130%;'>
                <hr style='border: 2px double rgb(0, 52, 120); '></hr>
                Chúng tôi xin cảm ơn thời gian và sự hợp tác của quý khách!
                <br/><br/>
                Trân Trọng,<br/>
                <b>Shah Ruchik</b><br/>
                Tổng Giám Đốc<br/>
                Ford Việt Nam<br/>
            <td>
        </tr>
        <tr>
            <td style='padding:30px 40px;'>
                <span style="color: black !important; font-size: 8pt; font-family: Arial, sans-serif; line-height: 150%;">
                    Lưu ý: Một số phần mềm e-mail có thể chia đường link trên thành hai dòng. Nếu mã số truy cập khảo sát của anh/chị không tự động điền vào màn hình khảo sát, xin anh/chị vui lòng cắt và dán thông tin sau vào ô mã số truy cập khảo sát TestID5045
                </span>
                <br/><br/>
                <span style="color: black !important; font-size: 8pt; font-family: Arial, sans-serif; line-height: 150%;">
                    Nếu anh/chị muốn xoá e-mail của mình ra khỏi danh sách của chúng tôi cho cuộc khảo sát này, xin anh/chị vui lòng nhấn vào <a href="https://ford-quality.com.vn/vg/Unsubscribe/Unsubscribe.cfm?surveyid=TestID5045" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" data-linkindex="1">đây</a> và nhập đia chị e-mail của mình. 
                </span>
                <br/><br/>
                <span style="color: black !important; font-size: 8pt; font-family: Arial, sans-serif; line-height: 150%;">
                    Dưới sự uỷ nhiệm của công ty Ford, bản nghiên cứu này được thực hiện bởi Ipsos, công ty nghiên cứu thị trường độc lập trụ sở tại US. Để xem chính sách riêng tư của Ipsos, xin anh/chị vui lòng nhấn vào <a href="https://ford-quality.com.vn/IpsosPrivacy" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" data-linkindex="2">đây</a>. Nếu anh/chị muốn tìm hiểu thêm về Ipsos, xin vui lòng truy cập <a href="https://www.Ipsos.com" target="_blank" rel="noopener noreferrer" data-auth="NotApplicable" data-linkindex="3">www.Ipsos.com</a>. 
                </span>
            </td>
        </tr>
    </table>
</center>
"""

#A string with the address of the sender
sender = "vn-survey@ipsosresearch.com"
#A list of string, one for each recipient
receivers = "long.pham@ipsos.com"

try:
    smtp_server = smtplib.SMTP('comxb.ipsosinteractive.com', 587) #smtp.office365.com
    smtp_server.ehlo()
    smtp_server.starttls()
    smtp_server.ehlo()
    smtp_server.login("vn-survey@ipsosresearch.com", "W6wx`Mx7*bJ6~5qF")
    smtp_server.ehlo()
    smtp_server.sendmail(sender, receivers, message.encode('utf-8').strip())
    smtp_server.quit()
    print("Successfully send email")
except smtplib.SMTPException as ex:
    print("Error: unable to send email - ", ex)
    

