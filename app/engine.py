
from __future__ import annotations
from dataclasses import dataclass
import pandas as pd
from .config import PRIMARY_WAREHOUSE
from .utils import normalize_token, safe_num, parse_dates

REQUIRED_CANONICAL = [
    "Fecha",
    "Código",
    "Tipo",
    "Bodega Origen",
    "Bodega Destino",
    "Descripción",
    "Cantidad",
    "Código Producto",
    "Nombre Producto",
    "Serie",
    "PVP",
    "Valor Unitario",
    "Valor Total",
    "Referencia",
    "Categoría Producto",
]

PHYSICAL_EXPECTED = ["Código Producto", "Nombre Producto", "Bodega", "Cantidad Física"]

ALIASES = {
    "fecha": "Fecha",
    "codigo": "Código",
    "codigo movimiento": "Código",
    "tipo": "Tipo",
    "bodega origen": "Bodega Origen",
    "bodega de origen": "Bodega Origen",
    "origen": "Bodega Origen",
    "bodega destino": "Bodega Destino",
    "bodega de destino": "Bodega Destino",
    "destino": "Bodega Destino",
    "descripcion": "Descripción",
    "detalle": "Descripción",
    "cantidad": "Cantidad",
    "cant": "Cantidad",
    "codigo producto": "Código Producto",
    "cod producto": "Código Producto",
    "sku": "Código Producto",
    "item": "Código Producto",
    "nombre producto": "Nombre Producto",
    "producto": "Nombre Producto",
    "descripcion producto": "Nombre Producto",
    "serie": "Serie",
    "pvp": "PVP",
    "valor unitario": "Valor Unitario",
    "v unitario": "Valor Unitario",
    "costo unitario": "Valor Unitario",
    "valor total": "Valor Total",
    "v total": "Valor Total",
    "referencia": "Referencia",
    "documento": "Referencia",
    "categoria producto": "Categoría Producto",
    "categoria": "Categoría Producto",
    "linea": "Categoría Producto",
    "bodega": "Bodega",
    "cantidad fisica": "Cantidad Física",
    "cantidad física": "Cantidad Física",
    # Variantes del archivo XLS del sistema contable (Contifico)
    "bodegadestino": "Bodega Destino",
    "codigo prod.": "Código Producto",
    "codigo prod": "Código Producto",
    "nombre prod.": "Nombre Producto",
    "nombre prod": "Nombre Producto",
    "categoria prod.": "Categoría Producto",
    "categoria prod": "Categoría Producto",
    "centro de costo": "Centro de Costo",
    "orden compra venta": "Orden Compra Venta",
}


@dataclass
class AnalysisResult:
    filtered: pd.DataFrame
    inventory_by_warehouse: pd.DataFrame
    sku_summary: pd.DataFrame
    samples_by_client: pd.DataFrame
    active_clients: pd.DataFrame
    kpis: dict
    physical_compare: pd.DataFrame | None
    warehouses: list[str]


class InventoryEngine:
    def __init__(self):
        self.raw_df: pd.DataFrame | None = None
        self.physical_df: pd.DataFrame | None = None
        self.excluded_skus: set[str] = set()
        # Bodegas excluidas globalmente: los movimientos donde la bodega
        # (origen o destino) esté aquí se eliminan ANTES del análisis.
        self.excluded_warehouses: set[str] = set()

    def _read_excel_flexible(self, path: str) -> pd.DataFrame:
        preview = pd.read_excel(path, header=None)
        best_row = 0
        best_score = -1

        for idx in range(min(10, len(preview))):
            row_values = [normalize_token(x) for x in preview.iloc[idx].tolist() if str(x).strip() and str(x) != "nan"]
            score = sum(1 for cell in row_values if cell in ALIASES)
            if score > best_score:
                best_score = score
                best_row = idx

        return pd.read_excel(path, header=best_row)

    def _canonicalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        ren = {}
        for col in df.columns:
            key = normalize_token(col)
            ren[col] = ALIASES.get(key, str(col).strip())
        df = df.rename(columns=ren)
        return df

    def _ensure_columns(self, df: pd.DataFrame, required: list[str]) -> None:
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Faltan columnas requeridas: {missing}")

    def load_inventory_file(self, path: str) -> pd.DataFrame:
        df = self._read_excel_flexible(path)
        df = self._canonicalize_columns(df)
        self._ensure_columns(df, REQUIRED_CANONICAL)

        for col in REQUIRED_CANONICAL:
            if col not in df.columns:
                df[col] = ""

        df["Fecha"] = parse_dates(df["Fecha"])
        df["Cantidad"] = safe_num(df["Cantidad"])
        df["Valor Unitario"] = safe_num(df["Valor Unitario"])
        df["Valor Total"] = safe_num(df["Valor Total"])

        text_cols = [
            "Código", "Tipo", "Bodega Origen", "Bodega Destino", "Descripción",
            "Código Producto", "Nombre Producto", "Serie", "Referencia", "Categoría Producto"
        ]
        for col in text_cols:
            df[col] = df[col].fillna("").astype(str).str.strip()

        self.raw_df = df
        return df

    def load_physical_count(self, path: str) -> pd.DataFrame:
        df = self._read_excel_flexible(path)
        df = self._canonicalize_columns(df)
        if "Bodega" not in df.columns:
            if "Bodega Destino" in df.columns:
                df["Bodega"] = df["Bodega Destino"]
            elif "Bodega Origen" in df.columns:
                df["Bodega"] = df["Bodega Origen"]

        self._ensure_columns(df, PHYSICAL_EXPECTED)
        df["Cantidad Física"] = safe_num(df["Cantidad Física"])
        for c in ["Código Producto", "Nombre Producto", "Bodega"]:
            df[c] = df[c].fillna("").astype(str).str.strip()
        self.physical_df = df[PHYSICAL_EXPECTED].copy()
        return self.physical_df

    @staticmethod
    def _starts(series: pd.Series, prefix: str) -> pd.Series:
        return series.fillna("").astype(str).str.upper().str.startswith(prefix.upper())

    def get_warehouses(self) -> list[str]:
        if self.raw_df is None:
            return []
        vals = set(self.raw_df["Bodega Origen"].dropna().astype(str).str.strip()) | \
               set(self.raw_df["Bodega Destino"].dropna().astype(str).str.strip())
        vals = {x for x in vals if x}
        return sorted(vals)

    def analyze(
        self,
        cutoff_date: str,
        warehouse_mode: str = "Todas",
        selected_warehouses: list[str] | None = None,
    ) -> AnalysisResult:
        if self.raw_df is None:
            raise ValueError("Primero cargue el archivo consolidado.")
        selected_warehouses = selected_warehouses or []

        df = self.raw_df.copy()
        df = df[df["Fecha"].notna()]
        df = df[df["Fecha"] <= pd.to_datetime(cutoff_date)]

        if self.excluded_skus:
            df = df[~df["Código Producto"].isin(self.excluded_skus)]

        # Exclusión global de bodegas: descartar movimientos donde origen
        # O destino estén en la lista (aplica a todo el pipeline: KPIs,
        # rotación, kardex, compras, etc.)
        if self.excluded_warehouses:
            _bad = self.excluded_warehouses
            df = df[
                (~df["Bodega Origen"].isin(_bad)) &
                (~df["Bodega Destino"].isin(_bad))
            ]

        if warehouse_mode == "Solo principal":
            df = df[
                (df["Bodega Origen"] == PRIMARY_WAREHOUSE) |
                (df["Bodega Destino"] == PRIMARY_WAREHOUSE)
            ]
        elif warehouse_mode == "Selección manual":
            if not selected_warehouses:
                raise ValueError("Seleccione al menos una bodega para el modo manual.")
            df = df[
                (df["Bodega Origen"].isin(selected_warehouses)) |
                (df["Bodega Destino"].isin(selected_warehouses))
            ]

        ref = df["Referencia"].fillna("").astype(str).str.upper()
        typ = df["Tipo"].fillna("").astype(str).str.upper()

        df["is_purchase"] = (typ == "ING") & ref.str.startswith("FAC")
        df["is_supplier_return"] = (typ == "EGR") & ref.str.startswith("NCT")
        df["is_sale"] = (typ == "EGR") & ref.str.startswith("FAC")
        df["is_customer_return"] = (typ == "ING") & ref.str.startswith("NCT")
        df["is_transfer"] = typ == "TRA"
        df["is_sample_out"] = (
            df["is_transfer"] &
            (df["Bodega Origen"] == PRIMARY_WAREHOUSE) &
            (df["Bodega Destino"] != PRIMARY_WAREHOUSE)
        )
        df["is_sample_in"] = (
            df["is_transfer"] &
            (df["Bodega Destino"] == PRIMARY_WAREHOUSE) &
            (df["Bodega Origen"] != PRIMARY_WAREHOUSE)
        )

        inventory_by_warehouse = self._inventory_by_warehouse(df)
        sku_summary = self._sku_summary(df, inventory_by_warehouse)
        samples_by_client = self._samples_by_client(df)
        active_clients = (
            samples_by_client[samples_by_client["Stock en Cliente"] > 0]
            .copy()
            .sort_values(["Stock en Cliente", "Cliente"], ascending=[False, True])
        )
        kpis = self._kpis(df, sku_summary)

        physical_compare = None
        if self.physical_df is not None:
            physical_compare = self._physical_compare(inventory_by_warehouse)
            if len(physical_compare):
                kpis["Exactitud inventario"] = float(
                    physical_compare["Coincide"].mean() * 100.0
                )

        return AnalysisResult(
            filtered=df,
            inventory_by_warehouse=inventory_by_warehouse,
            sku_summary=sku_summary,
            samples_by_client=samples_by_client,
            active_clients=active_clients,
            kpis=kpis,
            physical_compare=physical_compare,
            warehouses=self.get_warehouses(),
        )

    def _inventory_by_warehouse(self, df: pd.DataFrame) -> pd.DataFrame:
        rows = []

        for _, row in df.iterrows():
            sku = row["Código Producto"]
            name = row["Nombre Producto"]
            cat = row["Categoría Producto"]
            qty = float(row["Cantidad"])
            unit = float(row["Valor Unitario"])
            bo = row["Bodega Origen"]
            bd = row["Bodega Destino"]

            if row["is_purchase"]:
                rows.append([sku, name, cat, bd or PRIMARY_WAREHOUSE, qty, unit, "Compra"])
            elif row["is_supplier_return"]:
                rows.append([sku, name, cat, bo or PRIMARY_WAREHOUSE, -qty, unit, "Dev. Proveedor"])
            elif row["is_sale"]:
                rows.append([sku, name, cat, bo or PRIMARY_WAREHOUSE, -qty, unit, "Venta"])
            elif row["is_customer_return"]:
                rows.append([sku, name, cat, bd or PRIMARY_WAREHOUSE, qty, unit, "Dev. Cliente"])
            elif row["is_transfer"]:
                if bo:
                    label = "Muestra enviada" if row["is_sample_out"] else "Transferencia salida"
                    rows.append([sku, name, cat, bo, -qty, unit, label])
                if bd:
                    label = "Muestra devuelta" if row["is_sample_in"] else "Transferencia ingreso"
                    rows.append([sku, name, cat, bd, qty, unit, label])

        mov = pd.DataFrame(
            rows,
            columns=[
                "Código Producto", "Nombre Producto", "Categoría Producto",
                "Bodega", "Cantidad Neta", "Valor Unitario", "Grupo Movimiento"
            ]
        )
        if mov.empty:
            return pd.DataFrame(columns=[
                "Código Producto", "Nombre Producto", "Categoría Producto",
                "Bodega", "Stock", "Valor Unitario Promedio", "Valor Stock", "Grupo Visual"
            ])

        out = mov.groupby(
            ["Código Producto", "Nombre Producto", "Categoría Producto", "Bodega"],
            as_index=False
        ).agg(
            Stock=("Cantidad Neta", "sum"),
            Valor_Unitario_Promedio=("Valor Unitario", "mean")
        )
        out["Valor Stock"] = out["Stock"] * out["Valor_Unitario_Promedio"]
        out["Grupo Visual"] = out["Bodega"].apply(
            lambda x: "Disponible" if x == PRIMARY_WAREHOUSE else "Muestras / Otras Bodegas"
        )
        out = out.rename(columns={"Valor_Unitario_Promedio": "Valor Unitario Promedio"})

        # Eliminar filas con stock = 0
        out = out[out["Stock"].abs() > 0.0001]

        return out.sort_values(["Bodega", "Nombre Producto", "Código Producto"])

    def _sku_summary(self, df: pd.DataFrame, inventory: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame()

        summary = df.groupby(
            ["Código Producto", "Nombre Producto", "Categoría Producto"],
            as_index=False
        ).agg(
            Compras=("Cantidad", lambda s: s[df.loc[s.index, "is_purchase"]].sum()),
            Dev_Proveedor=("Cantidad", lambda s: s[df.loc[s.index, "is_supplier_return"]].sum()),
            Ventas=("Cantidad", lambda s: s[df.loc[s.index, "is_sale"]].sum()),
            Dev_Cliente=("Cantidad", lambda s: s[df.loc[s.index, "is_customer_return"]].sum()),
            Muestras_Enviadas=("Cantidad", lambda s: s[df.loc[s.index, "is_sample_out"]].sum()),
            Muestras_Devueltas=("Cantidad", lambda s: s[df.loc[s.index, "is_sample_in"]].sum()),
            Valor_Compras=("Valor Total", lambda s: s[df.loc[s.index, "is_purchase"]].sum()),
            Valor_Ventas=("Valor Total", lambda s: s[df.loc[s.index, "is_sale"]].sum()),
        )

        if inventory.empty:
            stock = pd.DataFrame(columns=[
                "Código Producto", "Stock Disponible", "Stock Muestras", "Stock Total", "Valor Inventario"
            ])
        else:
            inv = inventory.copy()
            inv["TipoStock"] = inv["Bodega"].apply(
                lambda x: "Stock Disponible" if x == PRIMARY_WAREHOUSE else "Stock Muestras"
            )
            pivot = inv.pivot_table(
                index="Código Producto",
                columns="TipoStock",
                values="Stock",
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            vals = inv.groupby("Código Producto", as_index=False)["Valor Stock"].sum().rename(
                columns={"Valor Stock": "Valor Inventario"}
            )
            stock = pivot.merge(vals, on="Código Producto", how="left")
            if "Stock Disponible" not in stock.columns:
                stock["Stock Disponible"] = 0.0
            if "Stock Muestras" not in stock.columns:
                stock["Stock Muestras"] = 0.0
            stock["Stock Total"] = stock["Stock Disponible"] + stock["Stock Muestras"]

        out = summary.merge(stock, on="Código Producto", how="left").fillna(0)
        paired_cols = [
            "Código Producto", "Nombre Producto", "Categoría Producto",
            "Compras", "Dev_Proveedor",
            "Ventas", "Dev_Cliente",
            "Muestras_Enviadas", "Muestras_Devueltas",
            "Stock Disponible", "Stock Muestras", "Stock Total", "Valor Inventario",
            "Valor_Compras", "Valor_Ventas",
        ]
        existing = [c for c in paired_cols if c in out.columns]
        return out[existing].sort_values(["Nombre Producto", "Código Producto"])

    def _samples_by_client(self, df: pd.DataFrame) -> pd.DataFrame:
        entreg = (
            df[df["is_sample_out"]]
            .groupby("Bodega Destino", as_index=False)["Cantidad"]
            .sum()
            .rename(columns={"Bodega Destino": "Cliente", "Cantidad": "Entregadas"})
        )
        dev = (
            df[df["is_sample_in"]]
            .groupby("Bodega Origen", as_index=False)["Cantidad"]
            .sum()
            .rename(columns={"Bodega Origen": "Cliente", "Cantidad": "Devueltas"})
        )
        out = entreg.merge(dev, on="Cliente", how="outer").fillna(0)
        if out.empty:
            return pd.DataFrame(columns=["Cliente", "Entregadas", "Devueltas", "Stock en Cliente"])
        out["Stock en Cliente"] = out["Entregadas"] - out["Devueltas"]
        return out.sort_values(["Stock en Cliente", "Cliente"], ascending=[False, True])

    def _kpis(self, df: pd.DataFrame, sku: pd.DataFrame) -> dict:
        stock_total = float(sku["Stock Total"].sum()) if not sku.empty and "Stock Total" in sku.columns else 0.0
        stock_disp = float(sku["Stock Disponible"].sum()) if not sku.empty and "Stock Disponible" in sku.columns else 0.0
        stock_muestras = float(sku["Stock Muestras"].sum()) if not sku.empty and "Stock Muestras" in sku.columns else 0.0
        valor_inv = float(sku["Valor Inventario"].sum()) if not sku.empty and "Valor Inventario" in sku.columns else 0.0
        ventas = float(df.loc[df["is_sale"], "Valor Total"].sum()) if not df.empty else 0.0
        compras = float(df.loc[df["is_purchase"], "Valor Total"].sum()) if not df.empty else 0.0

        if not df.empty:
            days = max(int((df["Fecha"].max() - df["Fecha"].min()).days) + 1, 1)
            consumo = float(df.loc[df["is_sale"], "Cantidad"].sum()) / days
            rot = (float(df.loc[df["is_sale"], "Cantidad"].sum()) / stock_total) if stock_total else 0.0
            dias_inv = (stock_total / consumo) if consumo else 0.0
        else:
            consumo = rot = dias_inv = 0.0

        margen = ((ventas - compras) / ventas * 100.0) if ventas else 0.0

        return {
            "Stock total": stock_total,
            "Stock disponible": stock_disp,
            "Stock en muestras": stock_muestras,
            "Valor inventario": valor_inv,
            "Ventas acumuladas": ventas,
            "Compras acumuladas": compras,
            "Rotación": rot,
            "Días de inventario": dias_inv,
            "Consumo promedio": consumo,
            "Margen": margen,
            "Exactitud inventario": 0.0,
        }

    def _physical_compare(self, inventory: pd.DataFrame) -> pd.DataFrame:
        calc = (
            inventory[["Código Producto", "Nombre Producto", "Bodega", "Stock"]]
            .rename(columns={"Stock": "Cantidad Calculada"})
            .copy()
        )
        phy = self.physical_df.copy()
        out = calc.merge(phy, on=["Código Producto", "Nombre Producto", "Bodega"], how="outer").fillna(0)
        out["Diferencia"] = out["Cantidad Física"] - out["Cantidad Calculada"]
        out["Coincide"] = out["Diferencia"].abs() < 0.0001
        return out.sort_values(["Bodega", "Nombre Producto", "Código Producto"])

    def export_result(self, result: AnalysisResult, output_path: str) -> None:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            result.filtered.to_excel(writer, sheet_name="Base Filtrada", index=False)
            result.inventory_by_warehouse.to_excel(writer, sheet_name="Inventario Bodega", index=False)
            result.sku_summary.to_excel(writer, sheet_name="Detalle SKU", index=False)
            result.samples_by_client.to_excel(writer, sheet_name="Muestras Cliente", index=False)
            result.active_clients.to_excel(writer, sheet_name="Clientes Activos", index=False)
            pd.DataFrame([result.kpis]).to_excel(writer, sheet_name="KPIs", index=False)
            if result.physical_compare is not None:
                result.physical_compare.to_excel(writer, sheet_name="Toma Fisica", index=False)
