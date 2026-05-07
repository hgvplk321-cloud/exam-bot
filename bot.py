import sys
import types
sys.modules['imghdr'] = types.ModuleType('imghdr')
import logging
import pandas as pd
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)

# ──────────────────────────────────────────────
# الإعدادات
# ──────────────────────────────────────────────
import os
BOT_TOKEN = os.getenv("BOT_TOKEN")غيّر هذا
EXCEL_FILE = "students.xlsx"

# حالات المحادثة
WAITING_ID = 1

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# تحميل بيانات الإكسل
# ──────────────────────────────────────────────
def load_data():
    """تحميل شيتات students و exam من ملف الإكسل"""
    xl = pd.ExcelFile(EXCEL_FILE)

    # شيت المتدربين
    students_df = xl.parse("students")
    students_df.columns = students_df.columns.str.strip()
    students_df["رقم المتدرب"] = students_df["رقم المتدرب"].astype(str).str.strip()
    students_df["الرقم المرجعي"] = students_df["الرقم المرجعي"].astype(str).str.strip()

    # شيت الاختبارات (الصف الأول فاضي، الثاني هو العناوين)
    exam_df = xl.parse("exam", header=1)
    exam_df.columns = exam_df.columns.str.strip()
    exam_df["الرقم المرجعي"] = exam_df["الرقم المرجعي"].astype(str).str.strip()

    return students_df, exam_df


try:
    students_df, exam_df = load_data()
    logger.info(f"✅ تم تحميل البيانات: {len(students_df)} متدرب | {len(exam_df)} اختبار")
except Exception as e:
    logger.error(f"❌ خطأ في تحميل الإكسل: {e}")
    raise


# ──────────────────────────────────────────────
# دالة البحث عن جدول الاختبارات
# ──────────────────────────────────────────────
def get_schedule(student_id: str) -> str:
    """
    1. البحث عن المقررات المسجلة للمتدرب
    2. البحث عن كل مقرر في جدول الاختبارات
    3. تجميع الرسالة
    """
    sid = str(student_id).strip()
    student_courses = students_df[students_df["رقم المتدرب"] == sid]

    if student_courses.empty:
        return "❌ لم يتم العثور على رقم تدريبي مطابق.\nتأكد من الرقم وأعد المحاولة."

    student_name = student_courses["اسم المتدرب"].iloc[0]
    lines = []
    lines.append(f"🎓 *اسم المتدرب:* {student_name}")
    lines.append(f"🔢 *الرقم التدريبي:* {sid}")
    lines.append("━━━━━━━━━━━━━━━━━━━━")
    lines.append("📋 *جدول الاختبارات النهائي:*\n")

    found_any = False

    for _, row in student_courses.iterrows():
        ref_num = str(row["الرقم المرجعي"]).strip()
        course_name = str(row.get("المقرر", "")).strip()
        exam_type = str(row.get("نوع الاختبار", "")).strip()
        student_status = str(row.get("حالة المتدرب", "")).strip()
        course_status = str(row.get("حالة المقرر", "")).strip()

        # البحث في جدول الاختبارات
        exam_rows = exam_df[exam_df["الرقم المرجعي"] == ref_num]

        if exam_rows.empty:
            # المقرر مسجل لكن ليس في جدول الاختبارات
            lines.append(f"📚 *المقرر:* {course_name}")
            lines.append(f"   🔖 الرقم المرجعي: {ref_num}")
            lines.append(f"   📝 نوع الاختبار: {exam_type}")
            lines.append(f"   👤 حالة المتدرب: {student_status}")
            lines.append(f"   📌 حالة المقرر: {course_status}")
            lines.append(f"   ⚠️ لا يوجد موعد في جدول الاختبارات")
            lines.append("")
        else:
            found_any = True
            for _, exam in exam_rows.iterrows():
                date = str(exam.get("التاريخ", "")).strip()
                day = str(exam.get("اليوم", "")).strip()
                period = str(exam.get("الفترة", "")).strip()
                location = str(exam.get("موقع اللجنة", "")).strip()
                committee = str(exam.get("اللجنة", "")).strip()

                # تجاهل صيغ XLOOKUP
                if location.startswith("="):
                    location = "—"

                lines.append(f"📚 *المقرر:* {course_name}")
                lines.append(f"   🔖 الرقم المرجعي: {ref_num}")
                lines.append(f"   📅 التاريخ: {date}  |  {day}")
                lines.append(f"   🕐 الفترة: {period}")
                lines.append(f"   📍 موقع اللجنة: {location}")
                lines.append(f"   🏷️ رقم اللجنة: {committee}")
                lines.append(f"   👤 حالة المتدرب: {student_status}")
                lines.append(f"   📌 حالة المقرر: {course_status}")
                lines.append("")

    if not lines[3:]:
        lines.append("ℹ️ لا توجد مقررات مسجلة لهذا الرقم.")

    lines.append("━━━━━━━━━━━━━━━━━━━━")
    lines.append("🏫 نظام جداول الاختبارات")

    return "\n".join(lines)


# ──────────────────────────────────────────────
# handlers المحادثة
# ──────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *أهلاً بك في نظام جداول الاختبارات!*\n\n"
        "أرسل /جدول للاستعلام عن جدول اختباراتك النهائي.",
        parse_mode="Markdown",
    )


async def request_schedule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🔢 *من فضلك، أدخل رقمك التدريبي:*",
        parse_mode="Markdown",
    )
    return WAITING_ID


async def receive_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    student_id = update.message.text.strip()

    # التحقق أن المدخل رقمي
    if not student_id.isdigit():
        await update.message.reply_text(
            "⚠️ الرقم التدريبي يجب أن يكون أرقاماً فقط.\nأعد الإدخال:"
        )
        return WAITING_ID

    await update.message.reply_text("⏳ جاري البحث...")

    result = get_schedule(student_id)
    await update.message.reply_text(result, parse_mode="Markdown")
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ تم إلغاء الطلب.")
    return ConversationHandler.END


async def fallback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """يستجيب للرسائل العامة"""
    text = update.message.text.strip()
    keywords = ["جدول", "اختبار", "جدولي", "مواعيد", "اختباراتي", "schedule"]
    if any(k in text for k in keywords):
        await update.message.reply_text(
            "🔢 *من فضلك، أدخل رقمك التدريبي:*",
            parse_mode="Markdown",
        )
        return WAITING_ID
    await update.message.reply_text(
        "أرسل /جدول للاستعلام عن جدول اختباراتك.\nأو أرسل /start للبداية."
    )
    return ConversationHandler.END


# ──────────────────────────────────────────────
# تشغيل البوت
# ──────────────────────────────────────────────
def main():
    app = ApplicationBuilder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("schedule", request_schedule),
            CommandHandler("schedule", request_schedule),
            MessageHandler(filters.TEXT & ~filters.COMMAND, fallback_handler),
        ],
        states={
            WAITING_ID: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_id)
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(conv_handler)

    logger.info("🚀 البوت يعمل...")
    app.run_polling()


if __name__ == "__main__":
    main()
