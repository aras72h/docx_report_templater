var html = `<!DOCTYPE html>
<html dir="rtl" lang="en-US">
	<head>
		<style>
      @page {
        size: A4;
        margin: 0;
      }
      @font-face {
          font-family: IranSans;
          src: url("/bpms/api/engine/engine/default/config/welcome/resource/font-fa.ttf");
      }
      .report-title p {
       display: inline;
      }
      body.rtl {
        direction: rtl;
        font-family: "IranSans", "IBMPlexSans", "open_sansregular", Helvetica,
          Arial, Verdana, sans-serif;
        font-size: 10pt;
        -webkit-print-color-adjust: exact;
        color-adjust: exact;
        text-align: justify;
        height: 262mm;
        padding: 0;
        margin: 0;
      }
      .flex-column {
        display: flex;
        flex-direction: column;
      }
      .flex-row {
        display: flex;
        flex-direction: row;
      }
      .mr-30 {
        margin-right: 30px;
      }
      .ml-30 {
        margin-left: 30px;
      }
      header {
        position: fixed;
        z-index:1;
        height: 80px;
        width: 100%;
        text-align: center;
        background: white;
        color: black;
        font-size: 16px;
        top: 0px;
      }
      .header-logo .border-bottom {
        border-bottom: 3px solid #c00000;
      }
      .header-title {
        font-size: 11pt;
        color: #002060;
        padding-top: 50px;
        padding-right: 10px;
      }
      footer {
        position: fixed;
        height: 100px;
        width: 100%;
        background: white;
        color: black;
        font-size: 9pt;
        bottom: 0;
      }
      footer .footer-two {
        width: 100%;
      }
      footer .footer-two .border-top {
        width: 85%;
        border-top: 3px solid #c00000;
      }
      footer .footer-two .content {
        margin-right: 50px;
      }
      .footer {
        position: absolute;
        height: 100px;
        width: 100%;
        text-align: center;
        background: white;
        color: black;
        font-size: 9pt;
        bottom: 0;
        z-index:2;
      }
      .footer .footer-one {
        width: 100%;
      }
      .footer .footer-one .border-top {
        width: 85%;
        border-top: 3px solid #c00000;
      }
      .footer .footer-one .address-info {
        display: flex;
        flex: 4 0
      }
      .footer .footer-one .pattern {
        display: flex;
        flex: 1 0;
        background:#002060;
        color: white;
        width:140px;
        float:right;
        display: flex;
        flex-direction: column;
        padding: 10px 20px;
        background: linear-gradient(90deg, #ffffff 20px, transparent 1%) center,
          linear-gradient(#ffffff 20px, transparent 0) center, #002060;
        background-size: 22px 22px;
         height: 74px;
        margin-top: 3px;
      }
      .header {
        position: absolute;
        z-index:2;
        height: 100px;
        width: 100%;
        text-align: center;
        background: white;
        color: black;
        font-size: 16px;
        top: 0px;
      }
      .side {
        background:#002060;
        color: white;
        width:140px;
        float:right;
        padding: 10px 20px;
        background: linear-gradient(90deg, #ffffff 20px, transparent 1%) center,
          linear-gradient(#ffffff 20px, transparent 0) center, #002060;
        background-size: 22px 22px;
        height: 259mm;
        position: relative;
        z-index: 3;
        margin-top: -40px;
        margin-left: 5px;
      }
      .report-content-section {
        position: relative;
        z-index: 3;
        margin-top: -80px;
      }
      .empty-header, .empty-footer {
        height:100px
      }
      .content {
        text-align: justify;
        /* margin-top: -120px; */
      }
      .items-info {
          background: white;
          align-self: center;
          align-items: center;
          justify-content: center;
          margin-bottom: 5px;
          text-align: center;
          width: 130px;
          font-size: 7pt;
      }
      .items-info .value {
        line-height: 1;
        margin: 0;
        word-break: break-all;
        max-height: 40px;
        overflow: hidden;
      }
      .networkArea-box {
        border-top: 5px solid #c00000;
        color: #1b3a7e;
      }
      .priority-box {
        color: white;
        border: 1px solid #c00000;
        border-radius: 5px;
        width: 60%;
        height: auto;
        background: #c00000;
        align-self: center;
        margin: 7px auto;
      }
      .items-box {
        background-color: white;
        border: 2px solid #1b3b7f;
        border-radius: 10px;
        color: black;
      }
      .items-box .title {
        border: 2px solid #1b3b7f;
        border-radius: 9px;
        color: white;
        background-color: #1b3b7f;
        width: 100%;
        text-align: center;
        height: 15px;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .text-center {
        text-align: center;
      }
      .date-area {
        padding: 0 50px;
        display: flex;
        justify-content: end;
      }
      .text {
        line-height: 2;
        font-weight: bold;
      }
      .p-20 {
        padding: 0 20px;
      }
      .p-40 {
        padding: 0 40px;
      }
      .font-size-11 {
        font-size: 11pt;
      }
      .align-items-center {
        align-items: center;
      }
      .align-items-end {
        align-items: end;
        text-align: center;
      }
      .arrow-left {
        font-size: 25pt;
        transform: rotate(180deg);
        margin-left: 10px;
      }
      .line-height-0 {
        line-height: 0;
      }
      .line-height-1 {
        line-height: 1;
      }
      .width-300 {
        width: 300px;
      }
      .square {
        border: 1px solid #c00000;
        background-color: #c00000;
        width: 12px;
        height: 12px;
        margin: 0 10px;
      }
      @media print {
        .header {
          background: white;
        }
        #pageNumber:after {
          counter-increment: page;
          content: counter(page);
        }
      }
    </style>
  </head>

	<body class="rtl"
  moznomarginboxes
  mozdisallowselectionprint>
  <div class="header"></div>
		<header class="headerTwo">
      <div class='flex-row'>
      <div class='header-nextPage'>
        <div class='header-logo'>
        <img height='80' width='100' src='https://mediamarketstreamer.iriborg.ir/vod/uploads/2023/8/22/8949f68e-e10c-4274-8ce6-1e475f89d813.png' /> 
        <div class='border-bottom'></div> </div> 
      </div>
      <div class='flex-row align-items-center header-title'>
        <p>گزارش نظارتی</p><p> ${data &&
        data.NetworkAreaName ? data.NetworkAreaName : ''} / صفحه  ... از ...</p> 
      </div>
    </div> 
    </header>
		<table>
			<thead><th><td><div class="empty-header"></div></td></th></thead>
			<tbody>
				<tr>
					<td>
						<div class="content">
							<div class="side">
                <div class="items-info">
                  <img height="120" width="100" src="https://mediamarketstreamer.iriborg.ir/vod/uploads/2022/2/5/a4793bf6-c21a-4eaa-82bd-b80bacd54fe5.png"/>
                </div>
                <div class="items-info networkArea-box">
                    <p>گزارش نظارتی</p>
                    <p>${data &&
        data.NetworkAreaName ? data.NetworkAreaName : ''}</p>
                </div>
                <div class="items-info priority-box">
                  <p>${data && data.NotePriorityId && data.PriorityTitle ? data.PriorityTitle : ""}</p>
                </div>
                <div class="items-info items-box flex-column">
                    <div class="title"><p>شبکه</p></div>
                    <div class="value"><p>${data &&
        data.NetworkName ? data.NetworkName : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p>برنامه</p></div>
                  <div class="value"><p>${data &&
        data.PTitle ? data.PTitle : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p>نوع گزارش</p></div>
                  <div class="value"><p>${data && data.IndicatorLeve1 && data.IndicatorLeve1Name ? data.IndicatorLeve1Name : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> تاریخ و زمان وقوع</p></div>
                  <div class="value" id="ReleaseDate">
                    <p><p>${data &&
        data.ReleaseDate ? `${new Date(data.ReleaseDate).toLocaleDateString("fa-IR")}` : ""}
                        </br>
                        ${data &&
        data.start
        ? `${data.end ?
            (function () {
                var seconds = 0;
                var date = new Date(1970, 0, 1);
                [data.end].forEach(element => {
                    var a = (element + '').split(":");
                    var sec = (+a[0]) * 60 * 60 + (+a[1]) * 60 + (+a[2]);
                    seconds = seconds + sec;
                });
                date.setSeconds(seconds);
                return `${date.toTimeString().replace(/.*(\d{2}:\d{2}:\d{2}).*/, "$1")}`
            })()
            : ""
        }
                            ${data.start ?
            (function () {
                var seconds = 0;
                var date = new Date(1970, 0, 1);
                [data.start].forEach(element => {
                    var a = (element + '').split(":");
                    var sec = (+a[0]) * 60 * 60 + (+a[1]) * 60 + (+a[2]);
                    seconds = seconds + sec;
                });
                date.setSeconds(seconds);
                return ` - ${date.toTimeString().replace(/.*(\d{2}:\d{2}:\d{2}).*/, "$1")}`
            })()
            : ""
        }
                            `
        : ""}
                          </p>
                  </div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> موضوع گزارش</p></div>
                  <div class="value"><p>${data && data.content ? data.content : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> نوع پخش</p></div>
                  <div class="value"><p>${data &&
        data.LiveBroadcastModeId && data.LiveBroadcastModeName ? data.LiveBroadcastModeName : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> قالب برنامه</p></div>
                  <div class="value"><p>${data &&
        data.FormatTypeTaggingId && data.FormatTypeTaggingName ? data.FormatTypeTaggingName : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> موضوع برنامه</p></div>
                  <div class="value"><p>${data && data.MetaSubject ? data.MetaSubject : ""}</p></div>
                </div>
                <div class="items-info items-box flex-column">
                  <div class="title"><p> کلید واژگان</p></div>
                  <div class="value"><p>${data && data.Keyword ? data.Keyword : ""}</p></div>
                </div>

              </div>
              <div class="report-content-section">
                <h3 class="text-center">بسمه تعالی</h3>
                <div class="date-area">
                  <div class="flex-column">
                    <div><label class="text">تاریخ : </label> ............</div>
                    <div><label class="text">شماره : </label> ............</div>
                    <div><label class="text">پیوست : </label> ............</div>
                  </div>
                </div>
                <div class="p-20">
                  <h2 class="font-size-11">${data.ReceiverPerson.AssociationPrintedReportSRName_Title} <br /> ${data.ReceiverPerson.AssociationPrintedReportSRPosition_Title} </h2>
                  <p class="font-size-11">سلام علیکم</p>
                  <div class="report-title ">
                    <p class=""> با احترام، گزارش نظارتی </p>
                    <p> با اهمیت
                      ${data && data.NotePriorityId ? data.PriorityTitle : ''} که حاصل بازبینی
                      برنامه های آن شبکه است، برای استحضار تقدیم می‌شود. </p>
                  </div>
                  <br/>
                  <div class="flex-row align-items-center line-height-0">
                    <span class="arrow-left">&#10147;</span>
                    <h3>روایت مسئله</h3>
                  </div>
                  <p>${data.TopicStory}</p>
  
                  <div class="flex-row align-items-center line-height-0">
                    <span class="arrow-left">&#10147;</span>
                    <h3>ملاحظه نظارتی / تحلیل کارشناسی</h3>
                  </div>
                  <p>${data.Considerations}</p>
                  <div class="flex-column align-items-end ml-30">
                    <h4 class="font-size-12 text-center">${data.SenderPerson.AssociationPrintedReportSRName_Title} <br />  ${data.SenderPerson.AssociationPrintedReportSRPosition_Title}</h4>
                  </div>
                  <div class="flex-column  width-300 mr-30">
                      <small class="line-height-1">رونوشت:</small>
                        <small>
                            ${data.Transcript
        ? data.Transcript.map(
            (item) =>
                `<br>${item.TranscriptTitle}</br>`
        ).join("")
        : ""}        
                        </small>
                  </div>
                </div>
              </div>
						</div>
					</td>
				</tr>
			</tbody>
			<tfoot>
				<tr>
				  <td><div class="empty-footer"></div></td>
				</tr>
			</tfoot>
		</table>
    <div class="footer">
      <div class="footer-one">
        <div class="border-top"></div>
        <div class="flex-row">
          <div class="pattern"></div>
          <div class="address-info flex-column">
            <p class="content">
              آدرس: تهران - خیابان بهشتی - بین خیابان شهید احمد قصیر و شهید خالد
              اسلامبولی - پلاک 285
            </p>
            <div class="flex-row align-items-center content">
              <div class='flex-row align-items-center'> 
                <div class='title'>کد پستی:</div> 
                <div class='value'> 1513617111</div> 
              </div> 
              <div class='square'></div> 
              <div class='flex-row align-items-center'> 
                <div class='title'>تلفن:</div> 
                <div class='value'> 88710193-4</div> 
              </div> 
              <div class='square'></div> 
              <div class='flex-row align-items-center'> 
                <div class='title'>دورنگار:</div> 
                <div class='value'> 88724557</div> 
              </div> 
              <div class='square'></div> 
              <div class='flex-row align-items-center'> 
                <div class='title'>رایانامه:</div> 
                <div class='value'> nezarat@irib.ir</div> 
              </div>
            </div>
          </div>
        </div>

      </div>
    </div>
		<footer>
      <div class="footer-two">
        <div class="border-top"></div>
        <p class="content"> 
          <b>مرکز نظارت و ارزیابی صداوسیما</b> : تهران - خیابان بهشتی - بین خیابان شهید احمد قصیر و شهید خالد اسلامبولی - پلاک 285 
        </p>
        <div class="flex-row align-items-center content">
          <div class='flex-row align-items-center'> 
            <div class='title'>کد پستی:</div> 
            <div class='value'> 1513617111</div> 
          </div> 
          <div class='square'></div> 
          <div class='flex-row align-items-center'> 
            <div class='title'>تلفن:</div> 
            <div class='value'> 88710193-4</div> 
          </div> 
          <div class='square'></div> 
          <div class='flex-row align-items-center'> 
            <div class='title'>دورنگار:</div> 
            <div class='value'> 88724557</div> 
          </div> 
          <div class='square'></div> 
          <div class='flex-row align-items-center'> 
            <div class='title'>رایانامه:</div> 
            <div class='value'> nezarat@irib.ir</div> 
          </div>
        </div>
      </div>
    </footer>
	
	</body>
</html>`;

console.log(data)

var frame1 = document.createElement('iframe');
frame1.name = "frame1";

document.body.appendChild(frame1);
var frameDoc = (frame1.contentWindow) ? frame1.contentWindow : (frame1.contentDocument.document) ? frame1.contentDocument.document : frame1.contentDocument;
frameDoc.document.open();
frameDoc.document.write(html);
frameDoc.document.close();
setTimeout(function () {
    window.frames["frame1"].focus();
    window.frames["frame1"].print();
    document.body.removeChild(frame1);
}, 500);