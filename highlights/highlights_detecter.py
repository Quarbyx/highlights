import docx
import json
document = docx.Document("file_ref.docx")
stored_lib = []
color_classification_map = {
    "FCFCFC": "Высшие корковые",
    "BDD6EE": "Зрительная",
    "C5E0B3": "Стволовая",
    "B4C6E7": "Пирамидная",
    "F7CAAC": "Мозжечковая",
    "FFF2CC": "Чувствительная",
    "DBDBDB": "Тазовые",
    "F2F2F2": "Амб. индекс.",
}
lib = json.dumps(stored_lib)
for paragraph in document.paragraphs[1:]: 
    existing_entry = None
    for run in paragraph.runs:
        r_element = run._element
        shading = r_element.find(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd")
        shading_val = shading.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill") if shading is not None else None
        if shading_val in color_classification_map:
                existing_entry = next((entry for entry in stored_lib if entry["классификация"] == color_classification_map[shading_val]), None)
                if existing_entry:
                    existing_entry["текст"] += " " + run.text
                else:
                    entry = {
                        "id": len(stored_lib) + 1,
                        "текст": run.text,
                        "классификация": color_classification_map[shading_val],
                        "EDSS": 0.0
                    }
                    stored_lib.append(entry)
print(stored_lib)
with open("marked.json", "w") as json_file:
    json.dump(stored_lib, json_file, ensure_ascii=False, indent=4)
