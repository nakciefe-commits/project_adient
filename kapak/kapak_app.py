import os
from PyQt6.QtWidgets import QMessageBox
from docxtpl import DocxTemplate
import shared.global_data as global_data

def generate_cover_report(main_window):
    try:
        seat_count = global_data.config.get("SEAT_COUNT", 1)
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        # Determine the template path based on seat count
        template_name = f"Kapak_{seat_count}.docx"
        template_path = os.path.join(root_dir, "kapak", template_name)
        
        if not os.path.exists(template_path):
            QMessageBox.warning(main_window, "Hata", f"Şablon bulunamadı: {template_name}\nLütfen 'kapak' klasöründe olduğundan emin olun.")
            return

        # Prepare context data for Word Template
        context = {
            "TEST_NAME": global_data.config.get("TEST_NAME", ""),
            "REPORT_NO": global_data.config.get("REPORT_NO", ""),
            "TEST_ID": global_data.config.get("TEST_ID", ""),
            "WO_NO": global_data.config.get("WO_NO", ""),
            "TEST_NO": global_data.config.get("TEST_NO", ""),
            "TEST_DATE": global_data.config.get("TEST_DATE", ""),
            "OEM": global_data.config.get("OEM", ""),
            "PROGRAM": global_data.config.get("PROGRAM", ""),
            "PURPOSE": global_data.config.get("PURPOSE", "")
        }

        # Add seat-specific data
        smp_ids = global_data.config.get("SMP_ID", ["", "", "", "", ""])
        test_samples = global_data.config.get("TEST_SAMPLE", ["", "", "", "", ""])

        for i in range(5):
            seat_idx = i + 1
            context[f"SMP_ID_{seat_idx}"] = smp_ids[i]
            context[f"TEST_SAMPLE_{seat_idx}"] = test_samples[i]

        # Render DocxTemplate
        doc = DocxTemplate(template_path)
        doc.render(context)
        
        # Save output to tempfiles
        tempfiles_dir = os.path.join(root_dir, "tempfiles")
        if not os.path.exists(tempfiles_dir):
            os.makedirs(tempfiles_dir)
            
        test_no_input = global_data.config.get("TEST_NO", "NoSet")
        suffix = test_no_input.split('/')[-1] if '/' in test_no_input else test_no_input
        out_filename = f"Kapak_{suffix}.docx"
        
        final_out_path = os.path.join(tempfiles_dir, out_filename)
        doc.save(final_out_path)
        
        QMessageBox.information(main_window, "Başarılı", f"Kapak raporu başarıyla oluşturuldu!\n\nDosya Yolu: {final_out_path}")

    except Exception as e:
        QMessageBox.critical(main_window, "Hata", f"Kapak oluşturulurken hata meydana geldi:\n{str(e)}")
