from __future__ import annotations

from datetime import datetime, timedelta


class PanelDataService:
    """Consulta única compartida por GUI, Excel y PDF del panel."""

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
    def _receipt_where(start_date, end_date, ars_filter, user_filter, medication, category):
        clauses = [
            "r.created_at IS NOT NULL",
            "r.created_at::timestamp::date BETWEEN %s::date AND %s::date",
            "r.is_deleted=0",
        ]
        params = [start_date, end_date]
        if ars_filter and ars_filter != "Todas las ARS":
            clauses.append("r.ars=%s")
            params.append(ars_filter)
        if user_filter and user_filter != "Todos los Usuarios":
            clauses.append("r.username=%s")
            params.append(user_filter)
        if medication and medication != "Todos los medicamentos":
            clauses.append(
                "EXISTS (SELECT 1 FROM recibo_items fx WHERE fx.recibo_id=r.id "
                "AND fx.categoria='Medicamentos' AND fx.nombre=%s)"
            )
            params.append(medication)
        if category and category != "Todas las categorías":
            clauses.append(
                "EXISTS (SELECT 1 FROM recibo_items fc WHERE fc.recibo_id=r.id AND fc.categoria=%s)"
            )
            params.append(category)
        return " AND ".join(clauses), params

    @staticmethod
    def _item_where(start_date, end_date, ars_filter, user_filter, medication, category):
        where, params = PanelDataService._receipt_where(
            start_date, end_date, ars_filter, user_filter, "", ""
        )
        clauses = [where]
        if medication and medication != "Todos los medicamentos":
            clauses.extend(["ri.categoria='Medicamentos'", "ri.nombre=%s"])
            params.append(medication)
        if category and category != "Todas las categorías":
            clauses.append("ri.categoria=%s")
            params.append(category)
        return " AND ".join(clauses), params

    def load(
        self,
        start_date: str,
        end_date: str,
        ars_filter: str = "Todas las ARS",
        user_filter: str = "Todos los Usuarios",
        medication: str = "Todos los medicamentos",
        category: str = "Todas las categorías",
        trend_granularity: str = "day",
    ):
        receipt_where, receipt_params = self._receipt_where(
            start_date, end_date, ars_filter, user_filter, medication, category
        )
        item_where, item_params = self._item_where(
            start_date, end_date, ars_filter, user_filter, medication, category
        )
        previous_start, previous_end = self._period_before(start_date, end_date)
        previous_where, previous_params = self._receipt_where(
            previous_start, previous_end, ars_filter, user_filter, medication, category
        )

        time_expression = {
            "day": "TO_CHAR(r.created_at::timestamp::date, 'YYYY-MM-DD')",
            "week": "TO_CHAR(DATE_TRUNC('week', r.created_at::timestamp), 'YYYY-MM-DD')",
            "month": "TO_CHAR(DATE_TRUNC('month', r.created_at::timestamp), 'YYYY-MM')",
        }.get(trend_granularity, "TO_CHAR(DATE_TRUNC('month', r.created_at::timestamp), 'YYYY-MM')")

        with self.connection_factory() as con:
            summary_row = con.execute(
                f"""SELECT COUNT(*) AS receipts, COALESCE(SUM(r.total), 0),
                           COALESCE(AVG(r.total), 0), COALESCE(SUM(r.sala), 0)
                    FROM recibos r WHERE {receipt_where}""",
                tuple(receipt_params),
            ).fetchone()
            summary = {
                "receipts": int(summary_row[0] or 0),
                "total": float(summary_row[1] or 0),
                "average": float(summary_row[2] or 0),
                "room": float(summary_row[3] or 0),
            }

            previous_row = con.execute(
                f"SELECT COUNT(*), COALESCE(SUM(r.total), 0) FROM recibos r WHERE {previous_where}",
                tuple(previous_params),
            ).fetchone()
            previous = {"receipts": int(previous_row[0] or 0), "total": float(previous_row[1] or 0)}

            trend_rows = con.execute(
                f"""SELECT {time_expression} AS period, COUNT(*) AS receipts,
                           COALESCE(SUM(r.total), 0) AS total
                    FROM recibos r WHERE {receipt_where}
                    GROUP BY 1 ORDER BY 1""",
                tuple(receipt_params),
            ).fetchall()
            trend = [
                {"label": str(row[0]), "receipts": int(row[1]), "total": float(row[2])}
                for row in trend_rows
            ]

            category_rows = con.execute(
                f"""SELECT ri.categoria, COUNT(DISTINCT r.id), COALESCE(SUM(ri.cantidad), 0),
                           COALESCE(SUM(ri.total), 0)
                    FROM recibo_items ri JOIN recibos r ON r.id=ri.recibo_id
                    WHERE {item_where}
                    GROUP BY ri.categoria ORDER BY SUM(ri.total) DESC""",
                tuple(item_params),
            ).fetchall()
            categories = [
                {"label": str(row[0]), "receipts": int(row[1]), "quantity": int(row[2]), "total": float(row[3])}
                for row in category_rows
            ]

            ars_rows = []
            if not ars_filter or ars_filter == "Todas las ARS":
                raw_ars_rows = con.execute(
                    f"""SELECT COALESCE(NULLIF(r.ars, ''), 'Sin ARS'), COUNT(*),
                               COALESCE(SUM(r.total), 0)
                        FROM recibos r WHERE {receipt_where}
                        GROUP BY COALESCE(NULLIF(r.ars, ''), 'Sin ARS')
                        ORDER BY SUM(r.total) DESC, COUNT(*) DESC""",
                    tuple(receipt_params),
                ).fetchall()
                ars_rows = [
                    {"label": str(row[0]), "receipts": int(row[1]), "total": float(row[2])}
                    for row in raw_ars_rows
                ]

            detailed_rows = con.execute(
                f"""SELECT r.numero, r.created_at, r.username, r.ars, ri.categoria,
                           ri.nombre, ri.cantidad, ri.precio_unit, ri.total
                    FROM recibo_items ri JOIN recibos r ON r.id=ri.recibo_id
                    WHERE {item_where}
                    ORDER BY r.created_at, r.numero, ri.categoria, ri.nombre""",
                tuple(item_params),
            ).fetchall()
            details = [
                {
                    "receipt": int(row[0]), "created_at": str(row[1] or ""),
                    "username": str(row[2] or ""), "ars": str(row[3] or ""),
                    "category": str(row[4]), "item": str(row[5]), "quantity": int(row[6]),
                    "unit_price": float(row[7]), "total": float(row[8]),
                }
                for row in detailed_rows
            ]

        for row in trend:
            row["average"] = row["total"] / row["receipts"] if row["receipts"] else 0.0
        for collection in (categories, ars_rows):
            for row in collection:
                row["average"] = row["total"] / row["receipts"] if row["receipts"] else 0.0

        distribution_total = sum(row["total"] for row in categories)
        for row in categories:
            row["percentage"] = row["total"] / distribution_total if distribution_total else 0.0

        distribution = [dict(row) for row in categories[:5]]
        if len(categories) > 5:
            remaining = categories[5:]
            other = {
                "label": "Otros",
                "receipts": sum(row["receipts"] for row in remaining),
                "quantity": sum(row["quantity"] for row in remaining),
                "total": sum(row["total"] for row in remaining),
            }
            other["average"] = other["total"] / other["receipts"] if other["receipts"] else 0.0
            other["percentage"] = other["total"] / distribution_total if distribution_total else 0.0
            distribution.append(other)

        ars_total = sum(row["total"] for row in ars_rows)
        ars_receipts = sum(row["receipts"] for row in ars_rows)
        summary_table = []
        for row in ars_rows:
            summary_table.append(
                {
                    "type": "ars",
                    "label": row["label"],
                    "receipts": row["receipts"],
                    "total": row["total"],
                    "average": row.get("average", 0.0),
                    "money_percentage": row["total"] / ars_total if ars_total else 0.0,
                    "receipt_percentage": row["receipts"] / ars_receipts if ars_receipts else 0.0,
                }
            )

        return {
            "start_date": start_date,
            "end_date": end_date,
            "filters": {
                "ars": ars_filter, "user": user_filter, "medication": medication,
                "category": category,
                "trend_granularity": trend_granularity,
            },
            "summary": summary,
            "previous": previous,
            "trend": trend,
            "monthly": trend,
            "categories": categories,
            "category_distribution": distribution,
            "users": [],
            "ars": ars_rows,
            "ars_breakdown": ars_rows,
            "show_ars_comparison": bool(not ars_filter or ars_filter == "Todas las ARS"),
            "bar": ars_rows,
            "comparison": ars_rows,
            "breakdown": ars_rows,
            "breakdown_type": "ars",
            "summary_table": summary_table,
            "details": details,
        }
