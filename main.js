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

// تبدیل تاریخ میلادی به شمسی
const dateObject = new Date("2024-01-15T12:00:00");
const options = { timeZone: 'Asia/Tehran', year: 'numeric', month: 'long', day: 'numeric' };
const formattedDate = dateObject.toLocaleString('fa-IR', options);


// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
    NetworkAreaName: "سیما", // this
    NetworkName: "شبکه یک",
    PTitle: "سلام صبح بخیر",
    IndicatorLevelName: "اولویت دیگران به خود", // this
    ReleaseDate: formattedDate,
    start: "08:00:00",
    end: "10:00:00",
    content: "نا هماهنگی", // ?
    LiveBroadcastModeName: "زنده",
    FormatTypeTaggingName: "ترکیبی-ترکیبی ساده",
    MetaSubject: "تاخیر", // this
    Keyword: "مشکل پخش",
    NotePriorityId: 1, // this
    PriorityTitle: "مهم", // ?
    // دریافت کننده
    AssociationPrintedReportSRName_Title: "احمد رضا ناظری", // جناب دارد؟
    AssociationPrintedReportSRPosition_Title: "مرکز پویش", // ؟؟
    // ارسال کننده همان اقدام‌کننده می‌شه؟
    AssociationPrintedReportSRName_Title: "فرهاد نجفی",
    AssociationPrintedReportSRPosition_Title: "مدیر کل نظارت سیما",
    TopicStory: "به گزارش خبرنگار سینمایی خبرگزاری فارس، هزینه بالای تولید فیلم در شرایط فعلی اقتصاد و فیلم‌سوزی سینما در سال‌های کرونا باعث شکست گیشه و هزاران میلیارد ضرر اهالی سینما شده بود. در حالی که این ورشکستگی گیشه هنر هفتم یکی از پیش‌بینی‌هایی بود که توسط اهالی سینما انجام می‌شد و وزیر ارشاد خبر از گیشه هزار میلیاردی سینمام تا پایان سال را داده بود،‌ مدیرعامل مؤسسه سینماشهر هاشم میرزاخانی گفت که « فروش سینما از ابتدای فروردین سال 1402 تا 24 دی به بیش از هزار میلیارد تومان رسید.» میزان این فروش بیش از فروش سینما در 3 سال گذشته بوده است.",
    Considerations: "سال گذشته فروش سینمای ایران با وجود 50 فیلم اکران شده به 440 میلیارد تومان رسیده بود و سال قبل از آن نیز گیشه سینمای سال 1400 با اکران 44 فیلم تنها 158 میلیارد فروش داشت. در سینمای 1402 ایران تاکنون بیش از 23 میلیون و 472 هزار نفر در 650 هزار سئانس به تماشای فیلم‌ها در سینماها پرداخته‌اند. فروش فیلم‌های اکران شده در سال 1402 بیش از 955 میلیارد تومان با بیش از 22 میلیون و 551 هزار مخاطب بوده است. همچنین در آمار ارائه شده، فروش فیلم‌های سینمایی مانند «ملاقات خصوصی»، «بخارست»، «خط استوا» و... که در سال گذشته اکران شدند و اکران آن تا امسال ادامه داشت نیز محاسبه شده است که بیش از 33 میلیارد تومان مربوط به این فیلم‌ها هستند. همچنین میانگین قیمت بلیت نیز در سال 1402 حدود 42 هزار تومان به ازای هر نفر است.",
});

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