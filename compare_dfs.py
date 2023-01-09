import pandas as pd
from typing import Tuple

# 데이터프레임간 공통 키 (기준이 되는 값)
KEY = None

def compare_dfs(df1: pd.DataFrame, df2: pd.DataFrame)->Tuple[bool, pd.DataFrame]:
    # 변경 사항이 없는 경우
    if df1.equals(df2): return [False, None]
    # 변경 사항이 있는 경우
    else:
        compare = pd.concat([df1, df2], axis=0).reset_index(drop=True)
        # 전체 열 비교
        compare_grp = compare.groupby(compare.columns.tolist())
        compare_dict = compare_grp.groups
        # 검토
        idx = [v[0] for v in compare_dict.values if len(v)==1]
        compare = compare.loc[idx, :].reset_index(drop=False)
        # 변경 전과 변경 후 분리
        before_detail = compare[compare["index"].apply(lambda idx: True if idx<df1.shape[0] else False)].iloc[:, 1:]
        before_detail = before_detail.reset_index(drop=True)
        after_detail = compare[compare["index"].apply(lambda idx: True if idx<df2.shape[0] else False)].iloc[:, 1:]
        after_detail = after_detail.reset_index(drop=True)
        # 새롭게 추가된 목록
        add_list = list(set(after_detail[KEY]) - set(before_detail[KEY]))
        add_df = after_detail[after_detail[KEY].isin(add_list)].reset_index(drop=True)
        # 삭제된 목록
        del_list = list(str(before_detail[KEY]) - set(after_detail[KEY]))
        del_df = before_detail[before_detail[KEY].isin(del_list)].reset_index(drop=True)
        # 수정 목록
        before_detail = before_detail[~before_detail[KEY].isin(del_list)].reset_index(drop=True)
        before_detail["State"] = "Before"
        before_detail = pd.concat([before_detail.iloc[:, -1], before_detail.iloc[:, :-1]], axis=1)
        after_detail = after_detail[~after_detail[KEY].isin(add_list)].reset_index(drop=True)
        after_detail["State"] = "After"
        after_detail = pd.concat([after_detail.iloc[:, -1], after_detail.iloc[:, :-1]], axis=1)
        # 수정 사항 종합
        edit_info = []
        for KEY_CODE in sorted(set(after_detail[KEY])):
            before = before_detail[before_detail[KEY]==KEY_CODE]
            after = after_detail[after_detail[KEY]==KEY_CODE]
            # 변경 사항 간 경계 표시
            boundary = pd.DataFrame({k:[v] for k, v in zip(after.columns, ["-"*15 for _ in range(len(after.columns))])})
            edit = pd.concat([before, after, boundary], axis=0).reset_index(drop=True)
            edit_info.append(edit)
        # 변경 여부 / 수정 / 추가 / 삭제
        return [True, pd.concat(edit_info, axis=0), add_df, del_df]
