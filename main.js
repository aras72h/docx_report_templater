// PizZip is required because docx/pptx/xlsx files are all zipped files, and
// the PizZip library allows us to load the file in memory
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
    path.resolve(__dirname, "input.docx"),
    "binary"
);

// Unzip the content of the file
const zip = new PizZip(content);

// This will parse the template, and will throw an error if the template is
// invalid, for example, if the template is "{user" (no closing tag)
const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

// Data
const dataObject = {
    "NetworkAreaName": "معاونت صدا",
    "NetworkName": "شبکه جوان (صدا)",
    "ReleaseDate": "2024-01-15T00:00:00",
    "SenderPerson": {},
    "ReceiverPerson": {},
    "Transcript": [],
    "content": "بیان مطلب بی اساس مجری و ایجاد مقایسه ایران با کشورهای دیگر از لحاظ تورم",
    "SupervisorFullName": "آمنه سادات احمدی",
    "start": "22:11:32.247",
    "TopicStory": "<p>مجری(صادق علیخانی): یکی از دوستان من رفته بود به یه کشوری اسم اون کشورو نمیارم بحث سیاسی نشه، می‌گفتش من رفتم یه قهوه خوردم بعد مثلاً یه دلار دادم به اون صاحب کافه، می‌گفت این بنده خدا رنگ و روش پرید یه لحظه موندکه &nbsp;من باید چیکار کنم با این یه دلار... میگه رفت خلاصه تموم مغازه‌های اطراف پولاشونو جمع کرد تا باقیمانده پول اون یه فنجون قهوه رو بابت اون یه دلار رو بده</p>",
    "Considerations": "<p>بیان چنین مطالبی در زمان نزدیک به انتخابات صحیح نیست و موجب مقایسه ایران و کشورهای دیگر از لحاظ تورم و گرانی می شود ضمن اینکه در سرچ اینترنتی و تحقیق از افراد ساکن در کشورهای دیگر چنین مطلبی یافت و اثبات نشد.</p>",
    "PTitle": "موکب جوانان",
    "MetaSubject": "تورم",
    "Keyword": "تورم، فنجان قهوه، اغراق",
    "PriorityTitle": "اغماض",
    "FormatTypeTaggingName": "گفتار محور -کارشناسی ",
    "LiveBroadcastModeName": "زنده",
    "IndicatorLeve1Name": "محتوایی",
    "ManagerDescription": "",
    "end": "22:12:33.061",
    "NoteEndTime": 79953.061,
    "NotePriorityId": 4,
    "FormatTypeTaggingId": 11363,
    "LiveBroadcastModeId": 1,
    "IndicatorLeve1": 1,
    "isReviewd": true,
    "AssigneeVar": "GeneralManager",
    "PlaylistBodyURL": "https://mediamarketstreamer.iriborg.ir/timeshift/607/2024-01-15/playlist?contentType=jsonfile",
    "CreateDate": "2024-01-17T09:40:11",
    "Person1": 2,
    "group": "comment",
    "Supervisor": "09124218537",
    "GeneralManager": "abolhasani",
    "created": "2024-01-17T09:39:55",
    "Playdate1": "2024-01-15T00:00:00",
    "MomentaryReviewedProcessId": 1006438,
    "BodyURL": "https://mediamarketstreamer.iriborg.ir/timeshift/607/2024-01-15/",
    "Result3": "3",
    "ModifiedRecordByUsr": "09122045764",
    "ModifyDate": "2024-01-17T10:11:55",
    "StreamNoteId": 1001053,
    "isCheck": true,
    "CreatedRecordByUsr": "09124218537"
};

// تبدیل تاریخ میلادی به شمسی
const dateObject = new Date(dataObject.ReleaseDate);
const options = { year: 'numeric', month: 'long', day: 'numeric' };
const formattedDate = dateObject.toLocaleString('fa-IR', options);

dataObject.ReleaseDate = formattedDate;

// حذف ثانیه و میلی‌ثانیه از زمان شروع و پایان
// const e2p = s => s.replace(/\d/g, d => '۰۱۲۳۴۵۶۷۸۹'[d]);

const startTime = dataObject.start.substring(0, 8);
dataObject.start = startTime;

const endTime = dataObject.end.substring(0, 8);
dataObject.end = endTime;


// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render(dataObject);

// Get the zip document and generate it as a nodebuffer
const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a
// file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, "output.docx"), buf);