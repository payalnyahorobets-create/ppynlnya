from __future__ import annotations

from pathlib import Path
import os

from flask import Flask, jsonify, render_template, request

from data_loader import ExcelStore, build_summary, default_excel_path


def create_app() -> Flask:
    app = Flask(__name__)

    excel_path = Path(os.getenv("EXCEL_PATH", "")).expanduser()
    if not excel_path.is_file():
        excel_path = default_excel_path()

    store = ExcelStore(excel_path) if excel_path.is_file() else None

    @app.route("/")
    def index() -> str:
        return render_template("index.html")

    @app.route("/api/summary")
    def summary() -> dict:
        if store is None:
            return jsonify(
                {
                    "error": "Excel file not found. Set EXCEL_PATH or place the workbook next to the app.",
                }
            ), 400
        info = build_summary(store)
        return jsonify(
            {
                "products_count": info.products_count,
                "analysis_count": info.analysis_count,
                "month_count": info.month_count,
                "month_sheets": info.month_sheets,
            }
        )

    @app.route("/api/products")
    def products() -> dict:
        if store is None:
            return jsonify({"error": "Excel file not found."}), 400
        limit = int(request.args.get("limit", 200))
        rows = store.head_as_records("Номенклатура", limit)
        columns = store.columns("Номенклатура")
        return jsonify({"columns": columns, "rows": rows})

    @app.route("/api/analysis")
    def analysis() -> dict:
        if store is None:
            return jsonify({"error": "Excel file not found."}), 400
        limit = int(request.args.get("limit", 200))
        rows = store.head_as_records("Аналіз ABC-XYZ", limit)
        columns = store.columns("Аналіз ABC-XYZ")
        return jsonify({"columns": columns, "rows": rows})

    @app.route("/api/monthly")
    def monthly() -> dict:
        if store is None:
            return jsonify({"error": "Excel file not found."}), 400
        sheet = request.args.get("sheet")
        if not sheet:
            return jsonify({"error": "sheet is required"}), 400
        available = set(store.month_sheets())
        if sheet not in available:
            return jsonify({"error": f"sheet not found: {sheet}"}), 400
        limit = int(request.args.get("limit", 200))
        rows = store.head_as_records(sheet, limit)
        columns = store.columns(sheet)
        return jsonify({"columns": columns, "rows": rows})

    return app


if __name__ == "__main__":
    application = create_app()
    application.run(host="0.0.0.0", port=5000, debug=False)
