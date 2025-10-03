#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
多個 XML -> 單一 CSV（主檔） + TSV（明細檔，可選）

主檔欄位（含你指定的欄位）：
- source_file                          ← 來源檔名（如不需要可自行移除）
- name
- description
- CreationTimeUTC
- IsManuallyCorrected
- TotalDistanceKm                      ← 依 ViaPoint 經緯度順序 Haversine 計算
- RouteInfo_IgnoringRestrictions
- RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_ImageName
- RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_StartMapId
- RouteInfo_ViaPoints_NumVia
- RouteInfo_ViaPoints_ViaPoint         ← 「完整」JSON（包含所有 ViaPoint 資訊）

明細 TSV（每個 ViaPoint 一列，避免主檔欄位過長）：
- source_file, placemark_index, placemark_name, seq
- Position, Lat, Lon
- GroupID, Segment, Heading, Type, LinkToGeom, Direction, TTSRemark, WorkType, MMRule, ManeuverID, ManeuverNumber, IsDeadEnd

使用方式：
1) 指定多個檔名或萬用字元（程式內會展開）：
   python xml_to_csv.py *.xml -o all_routes.csv --detail-tsv routes_detail.tsv

2) 用資料夾掃描：
   python xml_to_csv.py --dir "D:/Download/DSNY/Routes-Stage.2025-09-05T09_57_29" -o all_routes.csv --detail-tsv routes_detail.tsv
"""

import argparse
import csv
import glob
import json
import math
from pathlib import Path
import xml.etree.ElementTree as ET


# ---------------------------
# 基礎工具
# ---------------------------

def clean_text(s: str, collapse_space: bool = True) -> str:
    """將 XML 文字節點統一清理：去除 \r \n \t，首尾去空白，必要時把多空白壓成單一空白。"""
    if s is None:
        return ""
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = s.strip()
    if collapse_space:
        s = " ".join(s.split())
    return s


def parse_text(elem, path: str, default: str = "") -> str:
    node = elem.find(path)
    if node is None or node.text is None:
        return default
    return clean_text(node.text)


def parse_int(elem, path: str, default: int = 0) -> int:
    txt = parse_text(elem, path, "")
    try:
        return int(txt)
    except Exception:
        return default


def parse_float_safe(s: str):
    try:
        return float(s)
    except Exception:
        return None


def haversine_km(lat1, lon1, lat2, lon2) -> float:
    """球面距離（km）"""
    R = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlmb / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c


# ---------------------------
# 解析 RouteInfo / ViaPoints
# ---------------------------

def parse_via_points(routeinfo_elem):
    """
    回傳：
    - via_points_min：主檔要用的「完整 JSON」清單（包含所有 ViaPoint 資訊）
    - via_points_detail：明細檔用的完整欄位
    - positions：[(lat, lon), ...] 供距離計算
    """
    via_points_min = []
    via_points_detail = []
    positions = []

    vp_parent = routeinfo_elem.find("./ViaPoints")
    if vp_parent is None:
        return via_points_min, via_points_detail, positions

    seq = 0
    for vp in vp_parent.findall("./ViaPoint"):
        seq += 1

        # Position 格式： "lon, lat"
        pos_txt = parse_text(vp, "./Position", "")
        lat, lon = None, None
        if pos_txt:
            parts = [p.strip() for p in pos_txt.split(",")]
            if len(parts) == 2:
                lon = parse_float_safe(parts[0])
                lat = parse_float_safe(parts[1])
                if lat is not None and lon is not None:
                    positions.append((lat, lon))

        # 明細欄位（完整）
        detail = {
            "Position": pos_txt,
            "Lat": lat if lat is not None else "",
            "Lon": lon if lon is not None else "",
            "GroupID": parse_text(vp, "./GroupID", ""),
            "Segment": parse_text(vp, "./Segment", ""),
            "Heading": parse_text(vp, "./Heading", ""),
            "Type": parse_text(vp, "./Type", ""),
            "LinkToGeom": parse_text(vp, "./LinkToGeom", ""),
            "Direction": parse_text(vp, "./Direction", ""),
            "TTSRemark": parse_text(vp, "./TTSRemark", ""),
            "WorkType": parse_text(vp, "./WorkType", ""),
            "MMRule": parse_text(vp, "./MMRule", ""),
            "ManeuverID": parse_text(vp, "./ManeuverID", ""),
            "ManeuverNumber": parse_text(vp, "./ManeuverNumber", ""),
            "IsDeadEnd": parse_text(vp, "./IsDeadEnd", ""),
            "seq": seq,
        }
        via_points_detail.append(detail)

        # 主檔完整 JSON（包含所有 ViaPoint 資訊）
        via_points_min.append({
            "Position": pos_txt,
            "GroupID": detail["GroupID"],
            "Segment": detail["Segment"],
            "Heading": detail["Heading"],
            "Type": detail["Type"],
            "LinkToGeom": detail["LinkToGeom"],
            "Direction": detail["Direction"],
            "WorkType": detail["WorkType"],
            "MMRule": detail["MMRule"],
            "ManeuverID": detail["ManeuverID"],
            "ManeuverNumber": detail["ManeuverNumber"],
            "IsDeadEnd": detail["IsDeadEnd"],
        })

    return via_points_min, via_points_detail, positions


def total_distance_km(positions) -> float:
    if len(positions) <= 1:
        return 0.0
    total = 0.0
    for i in range(1, len(positions)):
        lat1, lon1 = positions[i - 1]
        lat2, lon2 = positions[i]
        total += haversine_km(lat1, lon1, lat2, lon2)
    return total


# ---------------------------
# 主流程：解析單一 XML → 產列
# ---------------------------

def process_single_xml(xml_path: Path):
    """回傳 (main_rows, detail_rows)"""
    main_rows = []
    detail_rows = []

    tree = ET.parse(xml_path)
    root = tree.getroot()

    placemarks = root.findall(".//Placemark")
    for idx, pm in enumerate(placemarks, start=1):
        name = parse_text(pm, "./name")
        description = parse_text(pm, "./description")
        creation_time = parse_text(pm, "./CreationTimeUTC")
        is_manually_corrected = parse_text(pm, "./IsManuallyCorrected")

        routeinfo = pm.find("./RouteInfo")
        if routeinfo is None:
            row = {
                "name": name,
                "description": description,
                "CreationTimeUTC": creation_time,
                "IsManuallyCorrected": is_manually_corrected,
                "TotalDistanceKm": 0.0,
                "RouteInfo_IgnoringRestrictions": "",
                "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_ImageName": "",
                "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_StartMapId": "",
                "RouteInfo_ViaPoints_NumVia": 0,
                "RouteInfo_ViaPoints_ViaPoint": "[]",
            }
            main_rows.append(row)
            continue

        ignoring = parse_text(routeinfo, "./IgnoringRestrictions")
        image_name = parse_text(routeinfo, "./MapCorrectionInfo/DatasetInfo/ImageInfo/ImageName")
        start_map_id = parse_text(routeinfo, "./MapCorrectionInfo/DatasetInfo/ImageInfo/StartMapId")

        num_via = parse_int(routeinfo, "./ViaPoints/NumVia", 0)
        via_min, via_detail, positions = parse_via_points(routeinfo)

        # 給主檔的 JSON（單行、移除控制字元）
        via_json = json.dumps(via_min, ensure_ascii=False)
        via_json = via_json.replace("\r", "").replace("\n", "")

        row = {
            "name": name,
            "description": description,
            "CreationTimeUTC": creation_time,
            "IsManuallyCorrected": is_manually_corrected,
            "TotalDistanceKm": round(total_distance_km(positions), 6),
            "RouteInfo_IgnoringRestrictions": ignoring,
            "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_ImageName": image_name,
            "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_StartMapId": start_map_id,
            "RouteInfo_ViaPoints_NumVia": num_via if num_via else len(via_detail),
            "RouteInfo_ViaPoints_ViaPoint": via_json,
        }
        main_rows.append(row)

        # 明細檔加上來源/placemark 識別
        for d in via_detail:
            detail_rows.append({
                "placemark_index": idx,
                "placemark_name": name,
                "seq": d["seq"],
                "Position": d["Position"],
                "Lat": d["Lat"],
                "Lon": d["Lon"],
                "GroupID": d["GroupID"],
                "Segment": d["Segment"],
                "Heading": d["Heading"],
                "Type": d["Type"],
                "LinkToGeom": d["LinkToGeom"],
                "Direction": d["Direction"],
                "TTSRemark": d["TTSRemark"],
                "WorkType": d["WorkType"],
                "MMRule": d["MMRule"],
                "ManeuverID": d["ManeuverID"],
                "ManeuverNumber": d["ManeuverNumber"],
                "IsDeadEnd": d["IsDeadEnd"],
            })

    return main_rows, detail_rows


# ---------------------------
# 進出檔處理
# ---------------------------

def write_main_csv(out_csv: Path, rows: list):
    fieldnames_main = [
        "name",
        "description",
        "CreationTimeUTC",
        "IsManuallyCorrected",
        "TotalDistanceKm",
        "RouteInfo_IgnoringRestrictions",
        "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_ImageName",
        "RouteInfo_MapCorrectionInfo_DatasetInfo_ImageInfo_StartMapId",
        "RouteInfo_ViaPoints_NumVia",
        "RouteInfo_ViaPoints_ViaPoint",
    ]
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    with out_csv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames_main,
            delimiter=",",
            quotechar='"',
            quoting=csv.QUOTE_ALL,    # ★ 全欄位加引號，最穩
            lineterminator="\r\n",
            escapechar="\\",
        )
        writer.writeheader()
        for r in rows:
            # 再保險：移除 JSON 內任何殘留控制字元
            v = r.get("RouteInfo_ViaPoints_ViaPoint")
            if isinstance(v, str):
                r["RouteInfo_ViaPoints_ViaPoint"] = v.replace("\r", "").replace("\n", "")
            writer.writerow(r)


def write_detail_tsv(out_tsv: Path, rows: list):
    if not rows:
        return
    fieldnames_detail = [
        "placemark_index", "placemark_name", "seq",
        "Position", "Lat", "Lon",
        "GroupID", "Segment", "Heading", "Type", "LinkToGeom",
        "Direction", "TTSRemark", "WorkType", "MMRule",
        "ManeuverID", "ManeuverNumber", "IsDeadEnd",
    ]
    out_tsv.parent.mkdir(parents=True, exist_ok=True)
    with out_tsv.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames_detail,
            delimiter="\t",   # ★ TSV，避免逗號干擾
            quotechar='"',
            quoting=csv.QUOTE_MINIMAL,
            lineterminator="\r\n",
        )
        writer.writeheader()
        writer.writerows(rows)


# ---------------------------
# CLI
# ---------------------------

def collect_input_files(args):
    files = []

    # 1) 明確列的路徑或萬用字元樣式
    for pattern in args.input:
        # 支援 Windows 萬用字元：由程式展開
        expanded = glob.glob(pattern)
        files.extend(expanded)

    # 2) 指定資料夾掃描
    if args.dir:
        base = Path(args.dir)
        if args.recursive:
            files.extend([str(p) for p in base.rglob(args.pattern)])
        else:
            files.extend([str(p) for p in base.glob(args.pattern)])

    # 去重、轉 Path
    uniq = []
    seen = set()
    for f in files:
        p = str(Path(f))
        if p not in seen:
            seen.add(p)
            uniq.append(Path(p))
    return uniq


def main():
    ap = argparse.ArgumentParser(description="Convert one or more XML files into a single CSV (main) and an optional TSV (detail).")
    ap.add_argument("input", nargs="*", type=str, default=[], help="輸入檔或萬用字元（如 *.xml），可多個")
    ap.add_argument("-o", "--output-csv", type=Path, required=True, help="輸出主檔 CSV 路徑（UTF-8 BOM）")
    ap.add_argument("--detail-tsv", type=Path, help="可選：輸出明細 TSV 路徑（每個 ViaPoint 一列）")
    ap.add_argument("--dir", type=str, help="可選：要掃描的資料夾路徑")
    ap.add_argument("--pattern", type=str, default="*.xml", help="搭配 --dir 的檔案樣式（預設 *.xml）")
    ap.add_argument("--recursive", action="store_true", help="搭配 --dir 遞迴掃描")
    args = ap.parse_args()

    input_files = collect_input_files(args)
    if not input_files:
        raise SystemExit("找不到任何輸入 XML（請確認檔名/萬用字元/資料夾）。")

    all_main = []
    all_detail = []

    for xml_path in input_files:
        try:
            main_rows, detail_rows = process_single_xml(Path(xml_path))
            all_main.extend(main_rows)
            all_detail.extend(detail_rows)
        except ET.ParseError as e:
            print(f"[警告] 解析失敗（跳過）：{xml_path} | {e}")
        except OSError as e:
            print(f"[警告] 檔案讀取失敗（跳過）：{xml_path} | {e}")

    if not all_main:
        raise SystemExit("沒有可寫入的主檔資料。")

    write_main_csv(args.output_csv, all_main)

    if args.detail_tsv:
        write_detail_tsv(args.detail_tsv, all_detail)

    print(f"✅ 已輸出主檔 CSV：{args.output_csv}")
    if args.detail_tsv:
        print(f"✅ 已輸出明細 TSV：{args.detail_tsv}")


if __name__ == "__main__":
    main()
