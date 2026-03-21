import pandas as pd
import numpy as np


class Excel_Pandas:
    """Класс для анализа медицинского оборудования."""

    STATUS_MAPPING = {
        "broken": "faulty",
        "error": "faulty",
        "faulty": "faulty",
        "needs_repair": "faulty",
        "ok": "operational",
        "op": "operational",
        "operational": "operational",
        "working": "operational",
        "maintenance": "maintenance_scheduled",
        "maint_sched": "maintenance_scheduled",
        "maintenance_scheduled": "maintenance_scheduled",
        "service_scheduled": "maintenance_scheduled",
        "planned": "planned_installation",
        "planned_installation": "planned_installation",
        "scheduled_install": "planned_installation",
        "to_install": "planned_installation",
    }

    def __init__(self, filename: str) -> None:
        """Загрузка данных и подготовка DataFrame."""

        self.df = pd.read_excel(filename)

        self.df["warranty_until"] = pd.to_datetime(
            self.df["warranty_until"], errors="coerce"
        )
        self.df["last_calibration_date"] = pd.to_datetime(
            self.df["last_calibration_date"], errors="coerce"
        )
        self.df["install_date"] = pd.to_datetime(
            self.df["install_date"], errors="coerce"
        )

        self.df["status"] = (
            self.df["status"]
            .astype(str)
            .str.lower()
            .str.strip()
            .map(self.STATUS_MAPPING)
            .fillna("unknown")
        )

    def warranty_filter(self) -> pd.DataFrame:
        """Отчёт по срокам гарантии."""

        df = self.df.copy()
        today = pd.Timestamp.today().normalize()

        df["delta"] = (df["warranty_until"] - today).dt.days

        conditions = [
            df["delta"] < 0,
            df["delta"] <= 31,
            df["delta"] <= 365,
            df["delta"] > 365,
        ]

        choices = [
            "Гарантия истекла",
            "Менее месяца",
            "Менее года",
            "Более года",
        ]

        df["warranty_status"] = np.select(
            conditions,
            choices,
            default="Нет данных",
        )

        return df

    def sort_by_problems(self) -> pd.DataFrame:
        """Клиники с наибольшим количеством проблем."""

        df = self.df.copy()

        df["total_problems"] = (
            df["issues_reported_12mo"].fillna(0)
            + df["failure_count_12mo"].fillna(0)
        )

        result = (
            df.groupby(
                ["clinic_id", "clinic_name", "city"]
            )
            .agg(
                total_problems=("total_problems", "sum"),
                devices_count=("device_id", "count"),
            )
            .reset_index()
        )

        result = result.rename(
            columns={"total_problems": "problem_score"}
        )

        return result.sort_values(
            by="problem_score",
            ascending=False,
        )

    def calibration(self) -> pd.DataFrame:
        """Отчёт по калибровке."""

        df = self.df.copy()
        today = pd.Timestamp.today().normalize()

        def get_status(row):
            if pd.isna(row["last_calibration_date"]):
                return "Нет данных"

            if (
                pd.notna(row["install_date"])
                and row["last_calibration_date"] < row["install_date"]
            ):
                return "Неправильная дата калибровки"

            if (today - row["last_calibration_date"]).days > 365:
                return "Требуется калибровка"

            return "Калибровка в норме"

        df["calibration_status"] = df.apply(get_status, axis=1)

        return df[
            [
                "device_id",
                "clinic_name",
                "model",
                "last_calibration_date",
                "calibration_status",
            ]
        ]

    def equipment_count(self) -> pd.DataFrame:
        """Сводная таблица по клиникам и моделям."""

        pivot_df = pd.pivot_table(
            self.df,
            values="device_id",
            index=["clinic_name", "city"],
            columns="model",
            aggfunc="count",
            fill_value=0,
        ).reset_index()

        return pivot_df

    def save(self) -> None:
        """Сохранение отчётов в Excel."""

        with pd.ExcelWriter("result.xlsx") as writer:
            self.warranty_filter().to_excel(
                writer,
                sheet_name="Гарантия",
                index=False,
            )
            self.sort_by_problems().to_excel(
                writer,
                sheet_name="Проблемы",
                index=False,
            )
            self.calibration().to_excel(
                writer,
                sheet_name="Калибровка",
                index=False,
            )
            self.equipment_count().to_excel(
                writer,
                sheet_name="Сводная",
                index=False,
            )


def main() -> None:
    """Функция запуска."""

    app = Excel_Pandas("medical_diagnostic_devices_10000.xlsx")
    app.save()
    
    print("Готово")


if __name__ == "__main__":
    main()
