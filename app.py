from __future__ import annotations

import io
import os
import re
from datetime import datetime
from typing import Dict, Iterable, Optional

from flask import Flask, render_template, request, send_file, abort
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell
from docx.shared import Inches, Pt
from PIL import Image


app = Flask(__name__)


PLACEHOLDERS = {
	"{ФИОБел}": "fio_be",
	"{ФИОРусс}": "fio_ru",
	"{годМакет}": "year",
	"{паспортНомер}": "passport_id",
}

# Позиционные маркеры по позициям (1..14) — используем для посимвольной подстановки
POSITION_MARKERS = [
	"①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭",
]

# Параметры A4 landscape (в дюймах, учитываем поля)
A4_WIDTH_IN = 11.69
A4_HEIGHT_IN = 8.27
DEFAULT_MARGIN_IN = 0.5  # если в шаблоне другие поля, запас по безопасности
SAFE_WIDTH_IN = A4_WIDTH_IN - 2 * DEFAULT_MARGIN_IN
SAFE_HEIGHT_IN = A4_HEIGHT_IN - 2 * DEFAULT_MARGIN_IN


def _iter_paragraphs_in_cell(cell: _Cell) -> Iterable[Paragraph]:
	for p in cell.paragraphs:
		yield p
	for table in cell.tables:
		for row in table.rows:
			for nested_cell in row.cells:
				yield from _iter_paragraphs_in_cell(nested_cell)


def _iter_all_paragraphs(doc: Document) -> Iterable[Paragraph]:
	for p in doc.paragraphs:
		yield p
	for t in doc.tables:
		for row in t.rows:
			for cell in row.cells:
				yield from _iter_paragraphs_in_cell(cell)


def _fill_position_markers(paragraph: Paragraph, value: str) -> None:
	"""Заполняет позиционные маркеры ①..⑭ соответствующими символами value."""
	if not value:
		return
	for r in paragraph.runs:
		text = r.text.strip()
		if text in POSITION_MARKERS:
			idx = POSITION_MARKERS.index(text)
			r.text = value[idx] if idx < len(value) else ""


def _scale_to_a4(img: Image.Image) -> tuple[Inches, Inches]:
	"""Вычисляет размер картинки в дюймах с сохранением пропорций под зону A4 landscape."""
	w, h = img.size
	if w <= 0 or h <= 0:
		return Inches(SAFE_WIDTH_IN), Inches(SAFE_HEIGHT_IN)
	ratio_w = SAFE_WIDTH_IN / float(w)
	ratio_h = SAFE_HEIGHT_IN / float(h)
	ratio = min(ratio_w, ratio_h)
	return Inches(w * ratio), Inches(h * ratio)


def _replace_image_paragraph(paragraph: Paragraph, image_bytes: bytes) -> None:
	"""Заменяет целиком параграф-картинку: очищает текст, вставляет изображение нужного размера."""
	try:
		img = Image.open(io.BytesIO(image_bytes))
	except Exception as exc:  # noqa: BLE001
		raise ValueError(f"Некорректное изображение: {exc}") from exc
	width_in, height_in = _scale_to_a4(img)
	# Чистим текст runs
	for r in list(paragraph.runs):
		r.text = ""
	# Вставляем картинку в этот параграф
	run = paragraph.add_run()
	buf = io.BytesIO()
	# Пересохраним в PNG для совместимости
	img.save(buf, format="PNG")
	buf.seek(0)
	run.add_picture(buf, width=width_in, height=height_in)


def replace_placeholders(doc: Document, mapping: Dict[str, str], images: Dict[str, bytes] | None = None) -> None:
	"""Заменяет плейсхолдеры, сохраняет форматирование runs, заполняет позиционные маркеры и вставляет картинки.
	"""
	keys = list(mapping.keys())
	images = images or {}

	def replace_in_paragraph(paragraph: Paragraph) -> None:
		runs = paragraph.runs
		if not runs:
			return
		orig_texts = [r.text for r in runs]
		full = "".join(orig_texts)
		# Если параграф содержит маркер-картинку — обрабатываем отдельно и выходим
		if "{картинка1}" in full and images.get("image1"):
			_replace_image_paragraph(paragraph, images["image1"])  # type: ignore[arg-type]
			return
		if "{картинка2}" in full and images.get("image2"):
			_replace_image_paragraph(paragraph, images["image2"])  # type: ignore[arg-type]
			return

		new_text = full
		for key in keys:
			value = mapping.get(key, "") or ""
			new_text = new_text.replace(key, value)
		if new_text != full:
			pos = 0
			for idx, r in enumerate(runs):
				length = len(orig_texts[idx])
				segment = new_text[pos:pos + length]
				r.text = segment
				pos += length
			if pos < len(new_text) and runs:
				runs[-1].text = runs[-1].text + new_text[pos:]
			for j in range(len(runs)):
				if sum(len(t) for t in orig_texts[:j + 1]) > len(new_text):
					for k in range(j + 1, len(runs)):
						runs[k].text = ""
					break
		passport = mapping.get("{паспортНомер}", "") or ""
		if passport:
			_fill_position_markers(paragraph, passport)

	for p in _iter_all_paragraphs(doc):
		replace_in_paragraph(p)


@app.route("/", methods=["GET"])  # type: ignore[misc]
def index():
	return render_template("form.html")


@app.route("/generate", methods=["POST"])  # type: ignore[misc]
def generate():
	fio_be = (request.form.get("fio_be") or "").strip()
	fio_ru = (request.form.get("fio_ru") or "").strip()
	year = (request.form.get("year") or "").strip()
	passport_id = (request.form.get("passport_id") or "").strip()
	image1 = request.files.get("image1")
	image2 = request.files.get("image2")

	if not fio_be or not fio_ru or not year:
		abort(400, description="Все поля обязательны")
	if not year.isdigit() or not (1900 <= int(year) <= 2100):
		abort(400, description="Год должен быть числом 1900-2100")
	if passport_id:
		if len(passport_id) > 14 or not re.match(r"^[0-9A-Za-zА-Яа-яЁё\-]+$", passport_id):
			abort(400, description="Номер паспорта: только буквы/цифры/-, до 14 символов")
		passport_id = passport_id.upper()

	template_path = os.path.join(app.root_path, "templates", "template.docx")
	if not os.path.exists(template_path):
		abort(500, description="Не найден шаблон template.docx в templates/")

	try:
		doc = Document(template_path)
	except Exception as exc:  # noqa: BLE001
		abort(500, description=f"Ошибка чтения DOCX: {exc}")

	mapping: Dict[str, str] = {
		"{ФИОБел}": fio_be,
		"{ФИОРусс}": fio_ru,
		"{годМакет}": year,
		"{паспортНомер}": passport_id,
	}
	for i in range(1, 15):
		mapping[f"{{паспортНомер{i}}}"] = passport_id[i - 1] if len(passport_id) >= i else ""

	images: Dict[str, bytes] = {}
	if image1 and image1.filename:
		images["image1"] = image1.read()
	if image2 and image2.filename:
		images["image2"] = image2.read()

	replace_placeholders(doc, mapping, images)

	output = io.BytesIO()
	try:
		doc.save(output)
	except Exception as exc:  # noqa: BLE001
		abort(500, description=f"Ошибка сохранения DOCX: {exc}")

	output.seek(0)
	filename = f"doc_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
	return send_file(
		output,
		as_attachment=True,
		download_name=filename,
		mimetype=(
			"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
		),
	)


if __name__ == "__main__":
	app.run(host="0.0.0.0", port=5000, debug=True)
