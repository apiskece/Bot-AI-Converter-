import pandas as pd
from telegram import Update, InputFile
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters, CallbackContext
import os

# Fungsi untuk mengonversi data ke format vCard
def convert_to_vcf(data):
    vcf_lines = []
    for index, row in data.iterrows():
        vcf_lines.append("BEGIN:VCARD")
        vcf_lines.append("VERSION:3.0")
        vcf_lines.append(f"N:{row['Name']}")
        vcf_lines.append(f"TEL:{row['Phone']}")
        vcf_lines.append(f"EMAIL:{row['Email']}")
        vcf_lines.append("END:VCARD")
    return "\n".join(vcf_lines)

# Fungsi untuk memproses file .xlsx
def process_xlsx(file_path):
    try:
        data = pd.read_excel(file_path)
        required_columns = ['Name', 'Phone', 'Email']
        
        # Validasi kolom
        if not all(col in data.columns for col in required_columns):
            return None, "File tidak sesuai format. Pastikan kolom 'Name', 'Phone', dan 'Email' tersedia."
        
        return data, None
    except Exception as e:
        return None, str(e)

# Fungsi untuk menangani pesan yang berisi file
def handle_document(update: Update, context: CallbackContext):
    file = update.message.document.get_file()
    file_path = f"./{update.message.document.file_id}.xlsx"
    file.download(file_path)

    data, error = process_xlsx(file_path)
    
    if error:
        update.message.reply_text(error)
        return

    vcf_data = convert_to_vcf(data)
    vcf_file_path = f"./{update.message.document.file_id}.vcf"
    
    with open(vcf_file_path, 'w') as vcf_file:
        vcf_file.write(vcf_data)

    with open(vcf_file_path, 'rb') as vcf_file:
        update.message.reply_document(document=InputFile(vcf_file, filename=vcf_file_path))

    # Hapus file sementara
    os.remove(file_path)
    os.remove(vcf_file_path)

# Fungsi untuk memulai bot
def start(update: Update, context: CallbackContext):
    update.message.reply_text("Selamat datang! Kirim file .xlsx yang berisi data kontak (Name, Phone, Email).")

def main():
    # Ganti 'YOUR_TOKEN' dengan token bot Anda
    updater = Updater("8063782699:AAHe6IE1f2LA4SYL7rBNwB6UpabDWqP_hDA", use_context=True)
    
    dp = updater.dispatcher
    dp.add_handler(CommandHandler("start", start))
    dp.add_handler(MessageHandler(Filters.document.mime_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"), handle_document))

    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
