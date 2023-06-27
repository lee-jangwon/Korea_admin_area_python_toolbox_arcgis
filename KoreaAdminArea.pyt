# -*- coding: utf-8 -*-

import arcpy
import glob
import os
import pandas as pd
import shutil


from datetime import datetime
from pathlib import Path


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "대한민국 행정경계 통합 도구"
        self.alias = "KoreanAdminAreaCompleteTool"

        # List of tool classes associated with this toolbox
        self.tools = [ZipsToGDB, CompleteCode, JoinRelations]


class JoinRelations(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "지역 관계 생성하기"
        self.description = "시도, 시군구, 읍면동, 리간에 관계를 정리한다."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        sido_feature = arcpy.Parameter(
            displayName="시도 행정경계 피처를 선택해주세요.",
            name="sido_feature",
            datatype="DEFeatureClass",
            direction="Input",
        )
        sgg_feature = arcpy.Parameter(
            displayName="시군구 행정경계 피처를 선택해주세요.",
            name="sgg_feature",
            datatype="DEFeatureClass",
            direction="Input",
        )
        emd_feature = arcpy.Parameter(
            displayName="읍면동(법정) 행정경계 피처를 선택해주세요.",
            name="emd_feature",
            datatype="DEFeatureClass",
            direction="Input",
        )
        gemd_feature = arcpy.Parameter(
            displayName="읍면동(행정) 행정경계 피처를 선택해주세요.",
            name="gemd_feature",
            datatype="DEFeatureClass",
            direction="Input",
        )
        ri_feature = arcpy.Parameter(
            displayName="리 행정경계 피처를 선택해주세요.",
            name="ri_feature",
            datatype="DEFeatureClass",
            direction="Input",
        )
        params = [
            sido_feature,
            sgg_feature,
            emd_feature,
            gemd_feature,
            ri_feature,
        ]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        sido_feature = Path(parameters[0].valueAsText)
        sgg_feature = Path(parameters[1].valueAsText)
        emd_feature = Path(parameters[2].valueAsText)
        gemd_feature = Path(parameters[3].valueAsText)
        ri_feature = Path(parameters[4].valueAsText)

        # process sgg
        if "SIDO_CD" not in arcpy.ListFields(str(sgg_feature)):
            arcpy.management.AddField(
                str(sgg_feature), "SIDO_CD", "TEXT", field_length=2, field_alias="시도 코드"
            )
            arcpy.management.CalculateField(str(sgg_feature), "SIDO_CD", "!SIG_CD![:2]")

        # join sido to sgg
        arcpy.AddMessage("Joining SGG")
        arcpy.management.JoinField(
            str(sgg_feature),
            "SIDO_CD",
            str(sido_feature),
            "CTPRVN_CD",
            fields=["CTP_ENG_NM", "CTP_KOR_NM"],
        )

        # process emd
        if "SIG_CD" not in arcpy.ListFields(str(emd_feature)):
            arcpy.management.AddField(
                str(emd_feature), "SGG_CD", "TEXT", field_length=5, field_alias="시군구 코드"
            )
            arcpy.management.CalculateField(str(emd_feature), "SGG_CD", "!EMD_CD![:5]")
        arcpy.AddMessage("Joining Emd")
        arcpy.management.JoinField(
            str(emd_feature),
            "SGG_CD",
            str(sgg_feature),
            "SIG_CD",
            fields=["SIG_ENG_NM", "SIG_KOR_NM", "CTP_ENG_NM", "CTP_KOR_NM"],
        )

        # process gemd
        if "SIG_CD" not in arcpy.ListFields(str(gemd_feature)):
            arcpy.management.AddField(
                str(gemd_feature),
                "SGG_CD",
                "TEXT",
                field_length=5,
                field_alias="시군구 코드",
            )
            arcpy.management.CalculateField(str(gemd_feature), "SGG_CD", "!EMD_CD![:5]")
        arcpy.AddMessage("Joining Gemd")
        arcpy.management.JoinField(
            str(gemd_feature),
            "SGG_CD",
            str(sgg_feature),
            "SIG_CD",
            fields=["SIG_ENG_NM", "SIG_KOR_NM", "CTP_ENG_NM", "CTP_KOR_NM"],
        )

        if "SIG_CD" not in arcpy.ListFields(str(ri_feature)):
            arcpy.management.AddField(
                str(ri_feature),
                "SGG_CD",
                "TEXT",
                field_length=5,
                field_alias="시군구 코드",
            )
            arcpy.management.CalculateField(str(ri_feature), "SGG_CD", "!LI_CD![:5]")
        arcpy.AddMessage("Joining Ri")
        arcpy.management.JoinField(
            str(ri_feature),
            "SGG_CD",
            str(sgg_feature),
            "SIG_CD",
            fields=["SIG_ENG_NM", "SIG_KOR_NM", "CTP_ENG_NM", "CTP_KOR_NM"],
        )

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


class CompleteCode(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "법정동-행정동 매칭하기"
        self.description = "법정동만 있는 레이어에 행정동 열을 생성하고 매칭해줍니다."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        source_gdb = arcpy.Parameter(
            displayName="행정경계 피처가 있는 GDB를 선택해주세요.",
            name="source_gdb",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input",
        )
        match_table = arcpy.Parameter(
            displayName="행정안정부에서 내려받은 jscode 엑셀 파일을 선택해주세요.",
            name="match_table",
            datatype="DEFile",
            parameterType="Required",
            direction="Input",
        )
        params = [source_gdb, match_table]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    # def processSido(self, sido: Path, feature: arcpy.Feature):
    #     df = pd.read_excel(sido, engine="openpyxl")
    #     df = df.loc[df["시군구명"].isnull()].loc[df["읍면동명"].isnull()]

    # def load_table(self,)

    def execute(self, parameters, messages):
        """The source code of the tool."""
        source_gdb = parameters[0].value
        arcpy.env.workspace = source_gdb

        # 매칭 테이블은 한 번만 만들어서 돌려 쓰자
        match_table = Path(parameters[1].valueAsText)
        match_df = (
            pd.read_excel(match_table, engine="openpyxl")
            .drop(
                [
                    "시도명",
                    "시군구명",
                    "읍면동명",
                    "동리명",
                    "생성일자",
                ],
                axis=1,
            )
            .drop_duplicates("법정동코드")
            .set_index("법정동코드")["행정동코드"]
        )

        for f in arcpy.ListFeatureClasses():
            if str(f) == "SIDO":
                sido_fields = ["BJ_CD", "HJD_CD"]
                current_fields = arcpy.ListFields(f)
                if "HJD_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "HJD_CD", "TEXT", field_length=10, field_alias="행정동 코드"
                    )
                if "BJ_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "BJ_CD", "TEXT", field_length=10, field_alias="법정 코드"
                    )
                    arcpy.management.CalculateField(
                        f, "BJ_CD", "''.join([!CTPRVN_CD!,'00000000'])"
                    )
                with arcpy.da.UpdateCursor(f, sido_fields) as sido_cursor:
                    for row in sido_cursor:
                        row[1] = match_df.get(int(row[0]))
                        sido_cursor.updateRow(row)
            elif str(f) == "SGG":
                # process sgg
                sgg_fields = ["BJ_CD", "HJD_CD"]
                current_fields = arcpy.ListFields(f)
                if "HJD_CD" not in current_fields:
                    arcpy.AddMessage("Adding HJD_CD for SGG")
                    arcpy.management.AddField(
                        f, "HJD_CD", "TEXT", field_length=10, field_alias="행정동 코드"
                    )
                if "BJ_CD" not in current_fields:
                    arcpy.AddMessage("Adding BJ_CD for SGG")
                    arcpy.management.AddField(
                        f, "BJ_CD", "TEXT", field_length=10, field_alias="법정 코드"
                    )
                    arcpy.management.CalculateField(
                        f, "BJ_CD", "''.join([!SIG_CD!,'00000'])"
                    )
                with arcpy.da.UpdateCursor(f, sgg_fields) as sgg_cursor:
                    for row in sgg_cursor:
                        row[1] = match_df.get(int(row[0]))
                        arcpy.AddMessage(f"{row[0]}, {row[1]}")
                        sgg_cursor.updateRow(row)
            elif str(f) == "EMD":
                # process emd
                emd_fields = ["BJ_CD", "HJD_CD"]
                current_fields = arcpy.ListFields(f)
                if "HJD_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "HJD_CD", "TEXT", field_length=10, field_alias="행정동 코드"
                    )
                if "BJ_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "BJ_CD", "TEXT", field_length=10, field_alias="법정 코드"
                    )
                    arcpy.management.CalculateField(
                        f, "BJ_CD", "''.join([!EMD_CD!,'00'])"
                    )
                with arcpy.da.UpdateCursor(f, emd_fields) as emd_cursor:
                    for row in emd_cursor:
                        row[1] = match_df.get(int(row[0]))
                        emd_cursor.updateRow(row)
            elif str(f) == "RI":
                # process ri
                ri_fields = ["BJ_CD", "HJD_CD"]
                current_fields = arcpy.ListFields(f)
                if "HJD_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "HJD_CD", "TEXT", field_length=10, field_alias="행정동 코드"
                    )
                if "BJ_CD" not in current_fields:
                    arcpy.management.AddField(
                        f, "BJ_CD", "TEXT", field_length=10, field_alias="법정 코드"
                    )
                    arcpy.management.CalculateField(f, "BJ_CD", "!LI_CD!")
                with arcpy.da.UpdateCursor(f, ri_fields) as ri_cursor:
                    for row in ri_cursor:
                        row[1] = match_df.get(int(row[0]))
                        ri_cursor.updateRow(row)

        # self.processSido(match_table)
        match_df = pd.read_excel(match_table)

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return


class ZipsToGDB(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "일괄 압축해제 및 파일 GDB생성하기"
        self.description = "주소기반산업지원서비스에서 내려받은 구역의 도형(.zip)을 한 번에 언패킹하여 파일 GDB로 정리합니다."
        self.canRunInBackground = False
        self.feature_types = {
            "EMD": "EMD",
            "LI": "RI",
            "RI": "RI",
            "SIG": "SGG",
            "CTPRVN": "SIDO",
            "GEMD": "GEMD",
            "BAS": "BAS",
            "MAKAREA": "MAKAREA",
        }

    def getParameterInfo(self):
        """Define parameter definitions"""
        # parameter0: 주소기반산업지원서비스에서 내려받은 파일을 한 곳에 정리한 파일 위치
        base_directory = arcpy.Parameter(
            displayName="도형(.zip)은 어느 폴더에 있나요? (폴더)",
            name="base_directory",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        # parameter1: 최종으로 사용하게 될 공간참조
        result_spatial_reference = arcpy.Parameter(
            displayName="어떤 좌표체계를 사용할까요? (좌표체계)",
            name="result_spatial_reference",
            datatype="GPSpatialReference",
            parameterType="Required",
            direction="Input",
        )

        # parameter2: 압출 파일을 모든 정리가 끝난 뒤 삭제할지
        delete_files = arcpy.Parameter(
            displayName="모든 작업이 끝나면 압축 파일을 삭제할까요? (예/아니요)",
            name="delete_files",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input",
        )

        #  parameter3: 최종 파일 GDB가 저장될 경로
        result_path = arcpy.Parameter(
            displayName="만들어질 파일 GDB를 어디에 저장할까요? (폴더)",
            name="result_path",
            datatype="DEFolder",
            parameterType="Required",
            direction="Input",
        )

        params = [
            base_directory,
            result_spatial_reference,
            delete_files,
            result_path,
        ]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def unzip(self, zipfile_path: Path, delete_afterward: bool) -> str:
        sido_name = str(zipfile_path.name.split("_")[-1].split(".")[0])
        arcpy.AddMessage(f"{sido_name} 폴더를 압축 해제하고 있습니다.")
        shutil.unpack_archive(
            zipfile_path,
            zipfile_path.parent / "unzipped" / sido_name,
            "zip",
        )
        # 만약 사용자가 압축파일 삭제하기를 선택할 경우
        if delete_afterward:
            # os.remove(zipfile_path)
            arcpy.AddMessage(f"{zipfile_path.name}을(를) 삭제하고 있습니다.")  # for debugging
        return zipfile_path.name

    def rename_template(self, template: str) -> str:
        if template not in self.feature_types.keys():
            raise ValueError(f"{template}은 처리할 수 없는 형식입니다.")
        return self.feature_types.get(template)

    def execute(self, parameters, messages):
        """The source code of the tool."""

        # 작업할 기본 변수 설정
        base_directory = Path(parameters[0].valueAsText.strip())
        # result_spatial_reference = parameters[1].value
        delete_afterward = parameters[2].value
        result_path = Path(parameters[3].valueAsText.strip())
        original_spatial_reference = arcpy.SpatialReference(5179)

        # 압축 해제
        for dir in glob.glob(str(base_directory / "*")):
            if Path(dir).suffix != ".zip":
                arcpy.AddMessage(f"[WARNING] {dir} is not a zipfile.")
                continue
            self.unzip(Path(dir), delete_afterward)

        # 파일 GDB 만들기
        outname = f"{datetime.now().strftime('KoreaAdmin_%y%m%d')}"
        arcpy.management.CreateFileGDB(
            out_folder_path=str(result_path), out_name=outname
        )
        result_gdb = result_path / (".".join([outname, "gdb"]))
        arcpy.AddMessage(f"{str(result_gdb)} Created")

        maks = []
        sidos = []
        emds = []
        gemds = []
        bases = []
        ris = []
        sggs = []

        # 파일 GDB에 들어갈 템플릿 만들기
        # 법정시도 CTPR
        # 법정시군구 SIG
        # 법정읍면동 EMD
        # 법정리 RI
        # 행정구역 GEMD
        # 기초구역 BAS

        # 압축 해제한 폴더에서 작업
        unzipped_dir = base_directory / "unzipped"

        for ctpr in glob.glob(str(unzipped_dir / "*")):
            # ctpr_name = Path(ctpr).name
            for shp in glob.glob(str(Path(ctpr) / "*/*.shp")):
                shp_type = self.rename_template(shp.split("_")[-1].split(".")[0])
                if shp_type == "BAS":
                    bases.append(shp)
                elif shp_type == "SIDO":
                    sidos.append(shp)
                elif shp_type == "SGG":
                    sggs.append(shp)
                elif shp_type == "EMD":
                    emds.append(shp)
                elif shp_type == "GEMD":
                    gemds.append(shp)
                elif shp_type == "MAKAREA":
                    maks.append(shp)
                elif shp_type == "RI":
                    ris.append(shp)
                else:
                    pass

        arcpy.AddMessage("지번이 부여되지 않은 영역(MAKAREA)을 만들고 있습니다...")
        arcpy.management.Merge(maks, str(result_gdb / "MAKAREA"))
        arcpy.DefineProjection_management(
            str(result_gdb / "MAKAREA"), original_spatial_reference
        )

        arcpy.AddMessage("기초구역(BAS)를 만들고 있습니다...")
        arcpy.management.Merge(bases, str(result_gdb / "BAS"))
        arcpy.DefineProjection_management(
            str(result_gdb / "BAS"), original_spatial_reference
        )

        arcpy.AddMessage("법정 시군구(SGG)를 만들고 있습니다...")
        arcpy.management.Merge(sggs, str(result_gdb / "SGG"))
        arcpy.DefineProjection_management(
            str(result_gdb / "SGG"), original_spatial_reference
        )

        arcpy.AddMessage("법정 읍면동(EMD)을 만들고 있습니다...")
        arcpy.management.Merge(emds, str(result_gdb / "EMD"))
        arcpy.DefineProjection_management(
            str(result_gdb / "EMD"), original_spatial_reference
        )
        # arcpy.management.Project(str(result_gdb / "EMD"), result_spatial_reference)

        arcpy.AddMessage("행정 읍면동(GEMD)을 만들고 있습니다...")
        arcpy.management.Merge(gemds, str(result_gdb / "GEMD"))
        arcpy.DefineProjection_management(
            str(result_gdb / "GEMD"), original_spatial_reference
        )
        # arcpy.management.Project(str(result_gdb / "GEMD"), result_spatial_reference)

        arcpy.AddMessage("법정 리(RI)를 만들고 있습니다...")
        arcpy.management.Merge(ris, str(result_gdb / "RI"))
        arcpy.DefineProjection_management(
            str(result_gdb / "RI"), original_spatial_reference
        )
        # arcpy.management.Project(str(result_gdb / "RI"), result_spatial_reference)

        arcpy.AddMessage("법정 시도(SIDO)를 만들고 있습니다...")
        arcpy.management.Merge(sidos, str(result_gdb / "SIDO"))
        arcpy.DefineProjection_management(
            str(result_gdb / "SIDO"), original_spatial_reference
        )
        # arcpy.management.Project(str(result_gdb / "SIDO"), result_spatial_reference)

        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
