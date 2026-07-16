from __future__ import annotations

from datetime import datetime, timedelta


class PanelDataService:
    """Fuente única de datos para GUI, Excel y PDF del panel estadístico."""

    COVERAGE_EXPR = (
        "COALESCE(NULLIF(r.tipo_cobertura, ''), "
        "CASE WHEN COALESCE(r.ars, '')='' THEN 'NO_ASEGURADO' ELSE 'ASEGURADO' END)"
    )

    def __init__(self, connection_factory):
        self.connection_factory = connection_factory

    @staticmethod
    def _period_before(start_date: str, end_date: str):
        start = datetime.strptime(start_date, "%Y-%m-%d")
        end = datetime.strptime(end_date, "%Y-%m-%d")
        days = max(1, (end - start).days + 1)
        previous_end = start - timedelta(days=1)
        previous_start = previous_end - timedelta(days=days - 1)
        return previous_start.strftime("%Y-%m-%d"), previous_end.strftime("%Y-%m-%d")

    @staticmethod
    def _selection_filter(column: str, selection, clauses: list, params: list):
        if isinstance(selection, dict):
            values = [str(value) for value in selection.get("values", []) if str(value).strip()]
            mode = selection.get("mode", "include")
            if values:
                placeholder = "%s"
                if mode in ("exclude", "excluir"):
                    clauses.append(f"NOT (COALESCE({column}, '') = ANY({placeholder}))")
                else:
                    clauses.append(f"COALESCE({column}, '') = ANY({placeholder})")
                params.append(values)
            return
        ignored = {"", "Todas las ARS", "Todos los Usuarios", None}
        if selection not in ignored:
            clauses.append(f"{column}=%s")
            params.append(selection)

    @classmethod
    def _receipt_where(
        cls, start_date, end_date, ars_filter, user_filter, medication, category, coverage
    ):
        clauses = [
            "r.created_at IS NOT NULL",
            "r.created_at::timestamp::date BETWEEN %s::date AND %s::date",
            "r.is_deleted=0",
        ]
        params = [start_date, end_date]
        cls._selection_filter("r.ars", ars_filter, clauses, params)
        cls._selection_filter("r.username", user_filter, clauses, params)
        if coverage == "Asegurados":
            clauses.append(f"{cls.COVERAGE_EXPR}='ASEGURADO'")
        elif coverage == "No asegurados":
            clauses.append(f"{cls.COVERAGE_EXPR}='NO_ASEGURADO'")
        if medication and medication != "Todos los medicamentos":
            clauses.append(
                "EXISTS (SELECT 1 FROM recibo_items fx WHERE fx.recibo_id=r.id "
                "AND fx.categoria='Medicamentos' AND fx.nombre=%s)"
            )
            params.append(medication)
        if category and category not in ("Todas las categorías", "Sala de Emergencia"):
            clauses.append(
                "EXISTS (SELECT 1 FROM recibo_items fc WHERE fc.recibo_id=r.id AND fc.categoria=%s)"
            )
            params.append(category)
        return " AND ".join(clauses), params

    @classmethod
    def _item_where(
        cls, start_date, end_date, ars_filter, user_filter, medication, category, coverage
    ):
        where, params = cls._receipt_where(
            start_date, end_date, ars_filter, user_filter, "", "", coverage
        )
        clauses = [where]
        if medication and medication != "Todos los medicamentos":
            clauses.extend(["ri.categoria='Medicamentos'", "ri.nombre=%s"])
            params.append(medication)
        if category == "Sala de Emergencia":
            clauses.append("1=0")
        elif category and category != "Todas las categorías":
            clauses.append("ri.categoria=%s")
            params.append(category)
        return " AND ".join(clauses), params

    @staticmethod
    def _summary(con, where, params):
        row = con.execute(
            f"""SELECT COUNT(*), COALESCE(SUM(r.total), 0),
                       COALESCE(AVG(r.total), 0), COALESCE(SUM(r.sala), 0)
                FROM recibos r WHERE {where}""",
            tuple(params),
        ).fetchone()
        return {
            "receipts": int(row[0] or 0),
            "total": float(row[1] or 0),
            "average": float(row[2] or 0),
            "room": float(row[3] or 0),
        }

    @staticmethod
    def _fill_daily_gaps(rows, start_date, end_date):
        indexed = {row["label"]: row for row in rows}
        cursor = datetime.strptime(start_date, "%Y-%m-%d")
        end = datetime.strptime(end_date, "%Y-%m-%d")
        result = []
        while cursor <= end:
            label = cursor.strftime("%Y-%m-%d")
            result.append(indexed.get(label, {"label": label, "receipts": 0, "total": 0.0, "average": 0.0}))
            cursor += timedelta(days=1)
        return result

    def load(
        self,
        start_date: str,
        end_date: str,
        ars_filter=None,
        user_filter=None,
        medication: str = "Todos los medicamentos",
        category: str = "Todas las categorías",
        trend_granularity: str = "day",
        coverage: str = "Todas",
        compare_previous: bool = False,
        previous_start: str = "",
        previous_end: str = "",
    ):
        ars_filter = ars_filter or {"mode": "include", "values": []}
        user_filter = user_filter or {"mode": "include", "values": []}
        receipt_where, receipt_params = self._receipt_where(
            start_date, end_date, ars_filter, user_filter, medication, category, coverage
        )
        item_where, item_params = self._item_where(
            start_date, end_date, ars_filter, user_filter, medication, category, coverage
        )
        if not previous_start or not previous_end:
            previous_start, previous_end = self._period_before(start_date, end_date)
        previous_where, previous_params = self._receipt_where(
            previous_start, previous_end, ars_filter, user_filter, medication, category, coverage
        )
        previous_item_where, previous_item_params = self._item_where(
            previous_start, previous_end, ars_filter, user_filter, medication, category, coverage
        )

        time_expression = {
            "day": "TO_CHAR(r.created_at::timestamp::date, 'YYYY-MM-DD')",
            "week": "TO_CHAR(DATE_TRUNC('week', r.created_at::timestamp), 'YYYY-MM-DD')",
            "month": "TO_CHAR(DATE_TRUNC('month', r.created_at::timestamp), 'YYYY-MM')",
        }.get(trend_granularity, "TO_CHAR(r.created_at::timestamp::date, 'YYYY-MM-DD')")

        with self.connection_factory() as con:
            summary = self._summary(con, receipt_where, receipt_params)
            previous = self._summary(con, previous_where, previous_params) if compare_previous else {}

            trend_rows = con.execute(
                f"""SELECT {time_expression}, COUNT(*), COALESCE(SUM(r.total), 0)
                    FROM recibos r WHERE {receipt_where} GROUP BY 1 ORDER BY 1""",
                tuple(receipt_params),
            ).fetchall()
            trend = [
                {"label": str(row[0]), "receipts": int(row[1]), "total": float(row[2])}
                for row in trend_rows
            ]
            for row in trend:
                row["average"] = row["total"] / row["receipts"] if row["receipts"] else 0.0
            if trend_granularity == "day":
                trend = self._fill_daily_gaps(trend, start_date, end_date)

            def category_data(where, params, room, receipt_count):
                rows = con.execute(
                    f"""SELECT ri.categoria, COUNT(DISTINCT r.id), COALESCE(SUM(ri.cantidad), 0),
                               COALESCE(SUM(ri.total), 0)
                        FROM recibo_items ri JOIN recibos r ON r.id=ri.recibo_id
                        WHERE {where} GROUP BY ri.categoria ORDER BY SUM(ri.total) DESC""",
                    tuple(params),
                ).fetchall()
                data = [
                    {"label": str(row[0]), "receipts": int(row[1]), "quantity": int(row[2]), "total": float(row[3])}
                    for row in rows
                ]
                if room > 0 and category in ("Todas las categorías", "Sala de Emergencia"):
                    data.append({"label": "Sala de Emergencia", "receipts": receipt_count, "quantity": receipt_count, "total": room})
                total = sum(row["total"] for row in data)
                for row in data:
                    row["average"] = row["total"] / row["receipts"] if row["receipts"] else 0.0
                    row["percentage"] = row["total"] / total if total else 0.0
                data.sort(key=lambda item: item["total"], reverse=True)
                return data

            categories = category_data(item_where, item_params, summary["room"], summary["receipts"])
            previous_categories = (
                category_data(
                    previous_item_where, previous_item_params,
                    previous.get("room", 0.0), previous.get("receipts", 0),
                )
                if compare_previous else []
            )

            insured_where = f"{receipt_where} AND {self.COVERAGE_EXPR}='ASEGURADO'"
            raw_ars_rows = [] if coverage == "No asegurados" else con.execute(
                f"""SELECT COALESCE(NULLIF(r.ars, ''), 'Sin ARS'), COUNT(*), COALESCE(SUM(r.total), 0)
                    FROM recibos r WHERE {insured_where}
                    GROUP BY COALESCE(NULLIF(r.ars, ''), 'Sin ARS')
                    ORDER BY SUM(r.total) DESC, COUNT(*) DESC""",
                tuple(receipt_params),
            ).fetchall()
            ars_rows = [
                {"label": str(row[0]), "receipts": int(row[1]), "total": float(row[2])}
                for row in raw_ars_rows
            ]

            coverage_rows = con.execute(
                f"""SELECT {self.COVERAGE_EXPR}, COUNT(*), COALESCE(SUM(r.total), 0),
                           COALESCE(AVG(r.total), 0), COALESCE(SUM(r.sala), 0)
                    FROM recibos r WHERE {receipt_where}
                    GROUP BY {self.COVERAGE_EXPR} ORDER BY 1""",
                tuple(receipt_params),
            ).fetchall()
            coverage_stats = [
                {"label": "Asegurados" if row[0] == "ASEGURADO" else "No asegurados",
                 "receipts": int(row[1]), "total": float(row[2]), "average": float(row[3]), "room": float(row[4])}
                for row in coverage_rows
            ]

            detailed_rows = con.execute(
                f"""SELECT r.numero, r.created_at, r.username, r.ars, {self.COVERAGE_EXPR},
                           ri.categoria, ri.nombre, ri.cantidad, ri.precio_unit, ri.total
                    FROM recibo_items ri JOIN recibos r ON r.id=ri.recibo_id
                    WHERE {item_where}
                    ORDER BY r.created_at, r.numero, ri.categoria, ri.nombre""",
                tuple(item_params),
            ).fetchall()
            details = [
                {"receipt": int(row[0]), "created_at": str(row[1] or ""), "username": str(row[2] or ""),
                 "ars": str(row[3] or ""), "coverage": str(row[4] or ""), "category": str(row[5]),
                 "item": str(row[6]), "quantity": int(row[7]), "unit_price": float(row[8]), "total": float(row[9])}
                for row in detailed_rows
            ]

        for row in ars_rows:
            row["average"] = row["total"] / row["receipts"] if row["receipts"] else 0.0
        distribution_total = sum(row["total"] for row in categories)
        distribution = [dict(row) for row in categories[:7]]
        ars_total = sum(row["total"] for row in ars_rows)
        ars_receipts = sum(row["receipts"] for row in ars_rows)
        summary_table = [
            {"type": "ars", "label": row["label"], "receipts": row["receipts"], "total": row["total"],
             "average": row["average"], "money_percentage": row["total"] / ars_total if ars_total else 0.0,
             "receipt_percentage": row["receipts"] / ars_receipts if ars_receipts else 0.0}
            for row in ars_rows
        ]
        current_category_map = {row["label"]: row for row in categories}
        previous_category_map = {row["label"]: row for row in previous_categories}
        category_labels = list(current_category_map)
        category_labels.extend(label for label in previous_category_map if label not in current_category_map)
        category_comparison = [
            {
                "label": label,
                "current": current_category_map.get(label, {}).get("total", 0.0),
                "previous": previous_category_map.get(label, {}).get("total", 0.0),
            }
            for label in category_labels
        ]

        return {
            "start_date": start_date, "end_date": end_date,
            "previous_start": previous_start, "previous_end": previous_end,
            "filters": {"ars": ars_filter, "user": user_filter, "medication": medication,
                        "category": category, "coverage": coverage, "trend_granularity": trend_granularity,
                        "compare_previous": compare_previous},
            "summary": summary, "previous": previous, "trend": trend, "monthly": trend,
            "categories": categories, "previous_categories": previous_categories,
            "category_comparison": category_comparison, "category_distribution": distribution,
            "coverage": coverage_stats, "users": [], "ars": ars_rows, "ars_breakdown": ars_rows,
            "show_ars_comparison": coverage != "No asegurados" and bool(ars_rows),
            "bar": ars_rows, "comparison": ars_rows, "breakdown": ars_rows, "breakdown_type": "ars",
            "summary_table": summary_table, "details": details,
            "distribution_total": distribution_total,
        }
