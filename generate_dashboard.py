# -*- coding: utf-8 -*-

import os
import sys

# 强制标准输出/错误输出使用 UTF-8，
# 避免 Windows 控制台默认 GBK 编码导致 ✓ 等字符打印报错。
if sys.version_info[0] >= 3:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
import re
import io
import base64
import json
import html as html_lib
import unicodedata
from datetime import datetime

from pyecharts.charts import Sunburst, Tab
from pyecharts.components import Image
from pyecharts import options as opts
from pyecharts.commons.utils import JsCode
from pyecharts.globals import CurrentConfig


# ============================================================
# 0. 韦恩图依赖
# ============================================================

try:
    import matplotlib.pyplot as plt
    from matplotlib_venn import venn2
except ImportError:
    print("缺少 matplotlib 或 matplotlib-venn。")
    print("请运行：pip install matplotlib matplotlib-venn")
    raise SystemExit(1)


# ============================================================
# 1. Pyecharts 静态资源
# ============================================================

CurrentConfig.ONLINE_HOST = "https://assets.pyecharts.org/assets/v5/"


# ============================================================
# 2. Excel 配置
#
# 文件路径采用「脚本所在目录 + 细菌库.xlsx」的相对定位方式，
# 这样无论项目被移动到哪个位置都能正常运行。
#
# 也可以通过命令行第 1 个参数手动指定 Excel 路径，例如：
#   python generate_dashboard.py "D:\备份\细菌库_20261001.xlsx"
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 动态日期（YYYYMMDD），用于生成文件名，例如「菌种分布图（20261019）.html」
DATE_STR = datetime.now().strftime("%Y%m%d")

if len(sys.argv) > 1:
    file_path = sys.argv[1]
else:
    file_path = os.path.join(BASE_DIR, "细菌库.xlsx")

sheet_adult = 'nasal microbiota-adult'
sheet_child = 'nasal microbiota-child'


print("正在读取 Excel 数据...")


try:
    df_adult_raw = pd.read_excel(
        file_path,
        sheet_name=sheet_adult
    )

    df_child_raw = pd.read_excel(
        file_path,
        sheet_name=sheet_child
    )

    print(
        f"成功读取数据："
        f"成人表 {len(df_adult_raw)} 行，"
        f"儿童表 {len(df_child_raw)} 行。"
    )

except Exception as e:
    print(f"读取 Excel 失败：{e}")
    print("请检查文件路径和 Sheet 名称。")
    raise SystemExit(1)


# ============================================================
# 3. 数据清洗
# ============================================================

required_cols = [
    '新菌序号',
    '细菌名称',
    '培养条件',
    '所属门',
    '所属科',
    '英文名称',
    '保藏数量（管）',
    '特征',
    '备注',
    '获取渠道'
]


def clean_text(text):
    """清洗 Excel 文本。"""

    if pd.isna(text):
        return ""

    text = str(text).strip()

    # --------------------------------------------------------
    # 关键：把各种 Unicode 空白字符统一为普通半角空格。
    #
    # 数据里存在 NO-BREAK SPACE（\xa0）等非标准空格，
    # 会导致同一个菌因空格类型不同而被误判为两种菌，
    # 例如：
    #   "Staphylococcus\xa0capitis" 与 "Staphylococcus capitis"
    # 本应是同一个菌。这里统一规范化，避免此类歧义。
    # --------------------------------------------------------
    text = "".join(
        " " if unicodedata.category(ch) == "Zs" else ch
        for ch in text
    )

    text = re.sub(
        r'[\n\r\t\'\"\\<>]',
        ' ',
        text
    )

    text = re.sub(
        r'\s+',
        ' ',
        text
    )

    return text.strip()


def normalize_name(value):
    """
    将菌名规范化为用于唯一判定的标准形式：

    1. 去除首尾空格；
    2. 折叠多余空白（含 \xa0 等 Unicode 空格）；
    3. 统一转为小写。

    这样 "Staphylococcus\xa0capitis" 和
    "Staphylococcus capitis" 会被判定为同一个菌，
    而 "Streptococcus symci" 和
    "Streptococcus shenyangsis" 仍会被判定为不同菌。
    """

    if pd.isna(value):
        return ""

    text = str(value).strip()

    text = "".join(
        " " if unicodedata.category(ch) == "Zs" else ch
        for ch in text
    )

    text = re.sub(r'\s+', ' ', text)

    return text.strip().lower()


def safe_html(value):
    """Tooltip HTML 转义。"""

    return html_lib.escape(
        str(value),
        quote=True
    )


def get_merge_key(row):
    """
    菌种唯一聚合键：

    以「英文名称」作为唯一判定标准——英文名不同即视为不同菌种。

    设计原因：
    中文名可能相同但英文名不同（例如「沈阳链球菌」
    同时对应 Streptococcus symci 和 Streptococcus shenyangsis，
    二者是不同菌株），若以中文名聚合会把不同的菌错误合并。

    规则：
    1. 优先使用英文名称的规范化形式作为唯一键；
    2. 若无英文名称，退回使用中文名称的规范化形式；
    3. 两者皆无，标记为「未知菌种」。

    规范化（normalize_name）会统一大小写与空白，
    因此 "Staphylococcus\xa0capitis" 与
    "Staphylococcus capitis" 会被判定为同一种菌，
    而 "Streptococcus symci" 与
    "Streptococcus shenyangsis" 仍为不同菌种。
    """

    cn = row['细菌名称']
    en = row['英文名称']

    en_key = normalize_name(en)
    if en_key:
        return en_key

    cn_key = normalize_name(cn)
    if cn_key:
        return cn_key

    return '未知菌种'


def process_dataframe(df, source_label):

    df = df.copy()

    # --------------------------------------------------------
    # 清洗列名
    # --------------------------------------------------------

    df.columns = [
        str(col).strip()
        for col in df.columns
    ]

    # --------------------------------------------------------
    # 兼容：
    # 细菌名称（打✓）
    # 细菌名称（出现标色）
    # 等类似列名
    # --------------------------------------------------------

    for col in list(df.columns):

        if (
            '细菌名称' in col
            and
            col != '细菌名称'
        ):

            df.rename(
                columns={
                    col: '细菌名称'
                },
                inplace=True
            )

    # --------------------------------------------------------
    # 检查必要字段
    # --------------------------------------------------------

    missing_cols = [
        col
        for col in required_cols
        if col not in df.columns
    ]

    if missing_cols:

        print(
            f"严重错误：{source_label} 表缺少："
            f"{missing_cols}"
        )

        raise SystemExit(1)

    # 获取渠道向下填充
    df['获取渠道'] = (
        df['获取渠道'].ffill()
    )

    # 空值处理
    df = df.fillna('')

    # 文本清洗
    for col in required_cols:

        df[col] = (
            df[col].apply(
                clean_text
            )
        )

    # 来源
    df['数据来源'] = source_label

    # 聚合键
    df['聚合键'] = df.apply(
        get_merge_key,
        axis=1
    )

    return df


print("正在进行数据预清洗与合并...")


df_adult = process_dataframe(
    df_adult_raw,
    'Adult'
)

df_child = process_dataframe(
    df_child_raw,
    'Child'
)


df = pd.concat(
    [
        df_adult,
        df_child
    ],
    ignore_index=True
)


# ============================================================
# 4. 成人 / 儿童 / 交集判定
# ============================================================

print("正在判定物种归属...")


venn_mapping = {}


for key, group in df.groupby('聚合键'):

    if key == '未知菌种':
        continue

    sources = set(
        group['数据来源']
    )

    if (
        'Adult' in sources
        and
        'Child' in sources
    ):

        venn_mapping[key] = (
            '交集 (Intersection)'
        )

    elif 'Child' in sources:

        venn_mapping[key] = (
            '儿童特有 (Child Only)'
        )

    else:

        venn_mapping[key] = (
            '成人特有 (Adult Only)'
        )


df['韦恩分类'] = (
    df['聚合键'].map(
        venn_mapping
    )
)


# ============================================================
# 5. 菌种数量统计
# ============================================================

def unique_species_count(df_subset):

    return int(
        df_subset.loc[
            df_subset['聚合键'] != '未知菌种',
            '聚合键'
        ].nunique()
    )


# 总体
total_species_count = (
    unique_species_count(
        df
    )
)


# 成人
adult_species_count = (
    unique_species_count(
        df[
            df['数据来源']
            ==
            'Adult'
        ]
    )
)


# 儿童
child_species_count = (
    unique_species_count(
        df[
            df['数据来源']
            ==
            'Child'
        ]
    )
)


mapping_values = list(
    venn_mapping.values()
)


adult_only_count = (
    mapping_values.count(
        '成人特有 (Adult Only)'
    )
)


child_only_count = (
    mapping_values.count(
        '儿童特有 (Child Only)'
    )
)


intersection_count = (
    mapping_values.count(
        '交集 (Intersection)'
    )
)


print("")
print("========================================")
print("菌种数量统计")
print("----------------------------------------")
print(f"总体菌种：       {total_species_count} 种")
print(f"全部成人：       {adult_species_count} 种")
print(f"全部儿童：       {child_species_count} 种")
print(f"成人特有：       {adult_only_count} 种")
print(f"儿童特有：       {child_only_count} 种")
print(f"成人儿童共有：   {intersection_count} 种")
print("========================================")
print("")


# ============================================================
# 6. 给每个菌种建立唯一搜索 ID
# ============================================================

valid_species_keys = [

    key

    for key in (
        df['聚合键'].drop_duplicates()
    )

    if key != '未知菌种'
]


species_search_id_map = {

    key: f"species_{i:05d}"

    for i, key in enumerate(
        valid_species_keys,
        start=1
    )
}


# ============================================================
# 7. 搜索数据库
# ============================================================

print("正在建立菌种搜索索引...")


def unique_valid_values(
    group,
    column
):

    values = []

    for value in group[column]:

        value = str(
            value
        ).strip()

        if (
            value
            and
            value != '/'
            and
            value not in values
        ):

            values.append(
                value
            )

    return values


search_records = []


for merge_key in valid_species_keys:

    group = df[
        df['聚合键']
        ==
        merge_key
    ]

    chinese_names = unique_valid_values(
        group,
        '细菌名称'
    )

    english_names = unique_valid_values(
        group,
        '英文名称'
    )

    strain_ids = unique_valid_values(
        group,
        '新菌序号'
    )

    phylums = unique_valid_values(
        group,
        '所属门'
    )

    families = unique_valid_values(
        group,
        '所属科'
    )

    chinese_name = (
        chinese_names[0]
        if chinese_names
        else
        '无'
    )

    english_name = (
        english_names[0]
        if english_names
        else
        'Unknown sp.'
    )

    phylum = (
        phylums[0]
        if phylums
        else
        '未知门'
    )

    family = (
        families[0]
        if families
        else
        '未知科'
    )

    if (
        english_name
        and
        english_name != 'Unknown sp.'
    ):

        genus = (
            english_name.split()[0]
        )

    else:

        genus = '未知属'

    sources = set(
        group['数据来源']
    )

    if (
        'Adult' in sources
        and
        'Child' in sources
    ):

        source_text = (
            'Adult + Child'
        )

    elif 'Adult' in sources:

        source_text = (
            'Adult'
        )

    else:

        source_text = (
            'Child'
        )

    venn_class = (
        venn_mapping.get(
            merge_key,
            ''
        )
    )

    search_records.append({

        "search_id":
            species_search_id_map[
                merge_key
            ],

        "merge_key":
            merge_key,

        "chinese_name":
            chinese_name,

        "english_name":
            english_name,

        "chinese_aliases":
            chinese_names,

        "english_aliases":
            english_names,

        "strain_ids":
            strain_ids,

        "phylum":
            phylum,

        "family":
            family,

        "genus":
            genus,

        "source":
            source_text,

        "venn_class":
            venn_class
    })


print(
    f"搜索索引完成，共 "
    f"{len(search_records)} 个菌种。"
)


# ============================================================
# 8. 生成韦恩图
# ============================================================

print("正在生成韦恩图...")


fig = plt.figure(
    figsize=(8, 6)
)


venn2(

    subsets=(
        adult_only_count,
        child_only_count,
        intersection_count
    ),

    set_labels=(
        'Adult',
        'Child'
    )
)


plt.title(
    "Venn Diagram of Bacterial Species\n"
    "Adult vs Child"
)


venn_buffer = io.BytesIO()


plt.savefig(

    venn_buffer,

    format='png',

    dpi=300,

    bbox_inches='tight',

    facecolor='white'
)


plt.close(fig)


venn_buffer.seek(0)


venn_bytes = (
    venn_buffer.getvalue()
)


# ------------------------------------------------------------
# Base64 嵌入 HTML
# ------------------------------------------------------------

venn_base64 = (
    base64.b64encode(
        venn_bytes
    ).decode(
        'utf-8'
    )
)


venn_data_uri = (
    "data:image/png;base64,"
    +
    venn_base64
)


# ============================================================
# 9. 旭日图配色
# ============================================================

academic_colors = [

    "#3B4992",
    "#EE0000",
    "#008B45",
    "#631879",
    "#008280",
    "#BB0021",
    "#5F559B",
    "#A20056",
    "#808180"
]


# ============================================================
# 10. Tooltip
# ============================================================

tooltip_formatter = JsCode(
    """
    function(params) {

        if (
            params.data &&
            params.data.custom_html
        ) {

            return params.data.custom_html;

        }

        if (params.data) {

            return (
                '<div style="' +
                'padding:8px;' +
                'font-family:Arial;' +
                'color:#333;' +
                '">' +

                '<b>' +
                params.data.name +
                '</b>' +

                '<br/>' +

                '<span style="color:#666;">' +
                '记录数：' +
                params.value +
                '</span>' +

                '</div>'
            );

        }

        return '';
    }
    """
)


# ============================================================
# 11. 下载 + 一键还原
# ============================================================

def create_toolbox(
    download_name
):

    return opts.ToolboxOpts(

        is_show=True,

        orient="horizontal",

        pos_right="2%",

        pos_top="2%",

        item_size=22,

        item_gap=16,

        feature={

            "saveAsImage": {

                "show": True,

                "type": "png",

                "name":
                    download_name,

                "title":
                    "下载当前图",

                "pixelRatio":
                    3,

                "backgroundColor":
                    "#FFFFFF",

                "excludeComponents": [
                    "toolbox"
                ]
            },

            "restore": {

                "show":
                    True,

                "title":
                    "一键还原"
            }
        }
    )


# ============================================================
# 12. 创建旭日图
# ============================================================

def create_sunburst(
    df_subset,
    title,
    download_name
):

    tree_dict = {}

    # ========================================================
    # 构造数据树
    # ========================================================

    for merge_key, df_species in (
        df_subset.groupby(
            '聚合键'
        )
    ):

        if merge_key == '未知菌种':
            continue

        # ----------------------------------------------------
        # 英文名称
        # ----------------------------------------------------

        valid_eng = unique_valid_values(
            df_species,
            '英文名称'
        )

        full_eng_name = (
            valid_eng[0]
            if valid_eng
            else
            "Unknown sp."
        )

        # ----------------------------------------------------
        # 门
        # ----------------------------------------------------

        valid_phylums = unique_valid_values(
            df_species,
            '所属门'
        )

        phylum = (
            valid_phylums[0]
            if valid_phylums
            else
            '未知门'
        )

        # ----------------------------------------------------
        # 科
        # ----------------------------------------------------

        valid_families = unique_valid_values(
            df_species,
            '所属科'
        )

        family = (
            valid_families[0]
            if valid_families
            else
            '未知科'
        )

        # ----------------------------------------------------
        # 中文名称
        # ----------------------------------------------------

        valid_chi = unique_valid_values(
            df_species,
            '细菌名称'
        )

        chinese_name = (
            valid_chi[0]
            if valid_chi
            else
            "无"
        )

        # ----------------------------------------------------
        # 属 + 种
        # ----------------------------------------------------

        if (
            full_eng_name
            !=
            "Unknown sp."
        ):

            name_parts = (
                full_eng_name.split()
            )

            genus = (
                name_parts[0]
            )

        else:

            name_parts = []

            genus = (
                "未知属"
            )

        if len(
            name_parts
        ) > 1:

            species_epithet = (
                " ".join(
                    name_parts[1:]
                )
            )

        else:

            species_epithet = (
                "sp."
            )

        # ====================================================
        # 菌株详细信息
        # ====================================================

        strain_ids = []

        channels_html = ""

        for _, row in (
            df_species.iterrows()
        ):

            sid = str(
                row['新菌序号']
            ).strip()

            if (
                sid
                and
                sid != '/'
                and
                sid not in strain_ids
            ):

                strain_ids.append(
                    sid
                )

            channel = (
                row['获取渠道']
                or
                "无"
            )

            amount = str(
                row['保藏数量（管）']
            ).strip()

            amount_str = (
                f"{amount} 管"
                if amount
                else
                "无"
            )

            condition = (
                row['培养条件']
                or
                "无"
            )

            feature = (
                row['特征']
                or
                "无"
            )

            note = (
                row['备注']
                or
                "无"
            )

            channels_html += f"""

            <div style="
                margin-top:10px;
                padding:10px;
                background-color:#F8F9FA;
                border-left:4px solid #008B45;
                border-radius:4px;
            ">

                <span style="
                    color:#008B45;
                    font-size:14px;
                ">
                    <b>
                        获取渠道：
                        {safe_html(channel)}
                    </b>
                </span>

                <br/>

                <span style="color:#555;">
                    <b>[ 保藏数量 ]</b>
                </span>

                {safe_html(amount_str)}

                <br/>

                <span style="color:#555;">
                    <b>[ 培养条件 ]</b>
                </span>

                {safe_html(condition)}

                <br/>

                <span style="color:#555;">
                    <b>[ 特征描述 ]</b>
                </span>

                {safe_html(feature)}

                <br/>

                <span style="color:#555;">
                    <b>[ 备 注 ]</b>
                </span>

                {safe_html(note)}

                <br/>

            </div>

            """

        strain_id_str = (
            ", ".join(
                strain_ids
            )
            if strain_ids
            else
            "无"
        )

        # ====================================================
        # Tooltip HTML
        # ====================================================

        full_html = f"""

        <div style="
            padding:12px;
            max-width:470px;
            white-space:normal;
            line-height:1.6;
            text-align:left;
            font-family:Arial;
            color:#333;
        ">

            <strong style="
                font-size:16px;
                color:#3C5488;
                font-style:italic;
            ">
                {safe_html(full_eng_name)}
            </strong>

            <br/>

            <div style="
                border-top:1px solid #D3D3D3;
                margin:8px 0;
            ">
            </div>

            <span style="color:#555;">
                <b>[ 中文名称 ]</b>
            </span>

            {safe_html(chinese_name)}

            <br/>

            <span style="color:#555;">
                <b>[ 新菌序号 ]</b>
            </span>

            {safe_html(strain_id_str)}

            <br/>

            {channels_html}

        </div>

        """

        # ====================================================
        # Species 节点
        # ====================================================

        species_node = {

            "name":
                str(
                    species_epithet
                ),

            "value":
                int(
                    len(
                        df_species
                    )
                ),

            "custom_html":
                full_html,

            "search_id":
                species_search_id_map.get(
                    merge_key,
                    ""
                ),

            "full_english_name":
                full_eng_name,

            "chinese_name":
                chinese_name,

            "merge_key":
                merge_key
        }

        # ====================================================
        # 门 -> 科 -> 属 -> 种
        # ====================================================

        if phylum not in tree_dict:

            tree_dict[
                phylum
            ] = {}

        if family not in tree_dict[
            phylum
        ]:

            tree_dict[
                phylum
            ][
                family
            ] = {}

        if genus not in tree_dict[
            phylum
        ][
            family
        ]:

            tree_dict[
                phylum
            ][
                family
            ][
                genus
            ] = []

        tree_dict[
            phylum
        ][
            family
        ][
            genus
        ].append(
            species_node
        )

    # ========================================================
    # Tree Data
    # ========================================================

    tree_data = []

    for phylum, fam_dict in (
        tree_dict.items()
    ):

        phy_node = {

            "name":
                phylum,

            "children":
                []
        }

        for family, gen_dict in (
            fam_dict.items()
        ):

            fam_node = {

                "name":
                    family,

                "children":
                    []
            }

            for genus, species_list in (
                gen_dict.items()
            ):

                fam_node[
                    "children"
                ].append({

                    "name":
                        genus,

                    "children":
                        species_list
                })

            phy_node[
                "children"
            ].append(
                fam_node
            )

        tree_data.append(
            phy_node
        )

    # --------------------------------------------------------
    # 空数据保护
    # --------------------------------------------------------

    if not tree_data:

        tree_data = [

            {
                "name":
                    "该分类下暂无数据",

                "value":
                    1
            }
        ]

    # ========================================================
    # Sunburst
    # ========================================================
    #
    # 注意：
    #
    # 这里不再向 set_global_opts()
    # 传 animation_opts。
    #
    # 避免：
    #
    # TypeError:
    # Chart.set_global_opts() got an unexpected
    # keyword argument 'animation_opts'
    #
    # ========================================================

    sunburst = (

        Sunburst(

            init_opts=opts.InitOpts(

                width="100%",

                height="85vh",

                bg_color="#FFFFFF"
            )
        )

        .set_colors(
            academic_colors
        )

        .add(

            series_name="",

            data_pair=tree_data,

            highlight_policy="ancestor",

            # ------------------------------------------------
            # 外圈稍微收缩
            #
            # 原：94%
            # 新：92%
            #
            # 给浏览器 Canvas 外缘抗锯齿
            # 留出一点安全区域。
            # ------------------------------------------------

            radius=[
                "5%",
                "92%"
            ],

            sort_="desc",

            node_click="rootToNode",

            label_opts=opts.LabelOpts(

                is_show=True,

                formatter="{b}",

                font_size=10,

                color="#222",

                font_family="Arial"
            ),

            # ------------------------------------------------
            # 边缘杂线优化
            #
            # 原：
            # border_width=0.8
            # opacity=0.88
            #
            # 半透明图层在动画过程中容易产生
            # 重叠细线。
            #
            # 改为完全不透明。
            # ------------------------------------------------

            itemstyle_opts=opts.ItemStyleOpts(

                border_width=0.4,

                border_color="#FFFFFF",

                opacity=1.0
            )
        )

        .set_global_opts(

            title_opts=opts.TitleOpts(

                title=title,

                pos_left="center",

                pos_top="1%",

                title_textstyle_opts=
                    opts.TextStyleOpts(

                        font_size=20,

                        font_family="Arial",

                        color="#333333"
                    )
            ),

            tooltip_opts=opts.TooltipOpts(

                is_show=True,

                formatter=
                    tooltip_formatter,

                border_width=1,

                border_color=
                    "#E0E0E0",

                background_color=
                    "rgba(252,252,252,0.98)",

                is_enterable=True,

                is_confine=True
            ),

            toolbox_opts=create_toolbox(
                download_name
            )
        )
    )

    # ========================================================
    # 明确保留 ECharts 原有动画
    #
    # 直接写入 options，
    # 不通过 animation_opts 封装，
    # 因此对 pyecharts 版本更加兼容。
    # ========================================================

    sunburst.options[
        "animation"
    ] = True

    return sunburst


# ============================================================
# 13. 创建 7 个标签页
# ============================================================

print("正在构建 7 个标签页...")


tab = Tab(
    page_title=
        "Leyao Lab - 菌种库分布图"
)


# ============================================================
# ① 总体分布
# ============================================================

sunburst_all = create_sunburst(

    df,

    f"菌种分布总览（共 {total_species_count} 种）",

    "01_菌种分布总览"
)


tab.add(

    sunburst_all,

    f"① 总体分布 ({total_species_count}种)"
)


# ============================================================
# ② 成人全部
# ============================================================

df_adult_all = df[
    df['数据来源']
    ==
    'Adult'
]


sunburst_adult_all = create_sunburst(

    df_adult_all,

    f"成人菌种分布图（共 {adult_species_count} 种）",

    "02_成人菌种分布"
)


tab.add(

    sunburst_adult_all,

    f"② 全部成人 ({adult_species_count}种)"
)


# ============================================================
# ③ 儿童全部
# ============================================================

df_child_all = df[
    df['数据来源']
    ==
    'Child'
]


sunburst_child_all = create_sunburst(

    df_child_all,

    f"儿童菌种分布图（共 {child_species_count} 种）",

    "03_儿童菌种分布"
)


tab.add(

    sunburst_child_all,

    f"③ 全部儿童 ({child_species_count}种)"
)


# ============================================================
# ④ 成人特有
# ============================================================

df_adult_only = df[
    df['韦恩分类']
    ==
    '成人特有 (Adult Only)'
]


sunburst_adult_only = create_sunburst(

    df_adult_only,

    f"成人特有菌种（共 {adult_only_count} 种）",

    "04_成人特有菌种"
)


tab.add(

    sunburst_adult_only,

    f"④ 仅成人 ({adult_only_count}种)"
)


# ============================================================
# ⑤ 儿童特有
# ============================================================

df_child_only = df[
    df['韦恩分类']
    ==
    '儿童特有 (Child Only)'
]


sunburst_child_only = create_sunburst(

    df_child_only,

    f"儿童特有菌种（共 {child_only_count} 种）",

    "05_儿童特有菌种"
)


tab.add(

    sunburst_child_only,

    f"⑤ 仅儿童 ({child_only_count}种)"
)


# ============================================================
# ⑥ 成人儿童共有
# ============================================================

df_intersection = df[
    df['韦恩分类']
    ==
    '交集 (Intersection)'
]


sunburst_intersection = create_sunburst(

    df_intersection,

    f"成人与儿童共有菌种（共 {intersection_count} 种）",

    "06_成人儿童共有菌种"
)


tab.add(

    sunburst_intersection,

    f"⑥ 共有交集 ({intersection_count}种)"
)


# ============================================================
# ⑦ 韦恩图
# ============================================================

venn_image = Image()


venn_image.add(

    src=venn_data_uri,

    style_opts={

        "style": (
            "display:block;"
            "width:auto;"
            "max-width:90vw;"
            "max-height:calc(100vh - 220px);"
            "height:auto;"
            "margin:18px auto 30px auto;"
            "object-fit:contain;"
        ),

        "alt":
            "Adult Child Venn Diagram"
    }
)


tab.add(

    venn_image,

    f"⑦ 韦恩图 ({total_species_count}种)"
)


# ============================================================
# 14. 输出基础 HTML
# ============================================================

output_file = (
    f"菌种分布图（{DATE_STR}）.html"
)


tab.render(
    output_file
)


# ============================================================
# 15. Chart ID
# ============================================================

chart_ids = [

    sunburst_all.chart_id,

    sunburst_adult_all.chart_id,

    sunburst_child_all.chart_id,

    sunburst_adult_only.chart_id,

    sunburst_child_only.chart_id,

    sunburst_intersection.chart_id
]


chart_targets = {

    "overview":
        sunburst_all.chart_id,

    "adult_all":
        sunburst_adult_all.chart_id,

    "child_all":
        sunburst_child_all.chart_id,

    "adult_only":
        sunburst_adult_only.chart_id,

    "child_only":
        sunburst_child_only.chart_id,

    "intersection":
        sunburst_intersection.chart_id
}


venn_chart_id = (
    venn_image.chart_id
)


# ============================================================
# 16. 搜索数据库转 JSON
# ============================================================

search_data_json = json.dumps(

    search_records,

    ensure_ascii=False

).replace(
    "</",
    "<\\/"
)


chart_ids_json = json.dumps(
    chart_ids,
    ensure_ascii=False
)


chart_targets_json = json.dumps(
    chart_targets,
    ensure_ascii=False
)


# ============================================================
# 17. 搜索框 HTML
#
# 已经去掉：
# “支持中文菌名、英文菌名和新菌序号”
# 这一行显示文字。
#
# status 元素仍然保留，
# 供 JS 内部使用，但 CSS 隐藏。
# ============================================================

search_html = """

<div
    id="species-search-panel"
    class="species-search-panel"
>

    <div
        class="species-search-main"
    >

        <div
            class="species-search-icon"
        >
            🔍
        </div>

        <input
            id="species-search-input"
            class="species-search-input"
            type="text"
            autocomplete="off"
            placeholder="搜索中文菌名 / 英文菌名 / 新菌序号"
        />

        <button
            id="species-search-button"
            class="species-search-button"
            type="button"
        >
            搜索
        </button>

        <button
            id="species-search-clear"
            class="species-search-clear"
            type="button"
        >
            清空
        </button>

    </div>

    <div
        id="species-search-status"
        class="species-search-status"
    >
    </div>

    <div
        id="species-search-results"
        class="species-search-results"
    >
    </div>

</div>

"""


# ============================================================
# 18. 读取 HTML
# ============================================================

with open(
    output_file,
    'r',
    encoding='utf-8'
) as f:

    html = f.read()


# ============================================================
# 19. 插入搜索栏
# ============================================================

tab_pattern = (
    r'(<div\s+class="tab">.*?</div>)'
)


html, inserted_count = re.subn(

    tab_pattern,

    r'\1'
    +
    search_html,

    html,

    count=1,

    flags=re.S
)


if inserted_count == 0:

    print(
        "警告：没有自动找到 Tab 导航区域。"
    )


# ============================================================
# 20. 响应式 CSS
# ============================================================

responsive_css = """

<style>

html,
body {

    width:100%;

    min-height:100%;

    margin:0;

    padding:0;

    box-sizing:border-box;

    overflow-x:hidden;

    background:#FFFFFF;

    font-family:
        Arial,
        "Microsoft YaHei",
        "PingFang SC",
        sans-serif;
}


* {

    box-sizing:border-box;
}


/* =========================================================
   顶部标签页
   ========================================================= */

.tab {

    width:100%;

    display:flex;

    flex-wrap:wrap;

    align-items:center;

    justify-content:center;

    gap:4px;

    position:sticky;

    top:0;

    z-index:99999;

    padding:6px;

    background:
        rgba(
            247,
            248,
            250,
            0.98
        );

    border:none;

    border-bottom:
        1px
        solid
        #E5E7EB;

    box-shadow:
        0
        2px
        8px
        rgba(
            0,
            0,
            0,
            0.05
        );
}


.tab button {

    float:none !important;

    flex:
        1
        1
        180px;

    min-width:150px;

    max-width:260px;

    padding:
        10px
        13px;

    border:none;

    border-radius:6px;

    background:#F3F4F6;

    color:#374151;

    font-size:
        clamp(
            12px,
            0.90vw,
            15px
        );

    white-space:nowrap;

    cursor:pointer;

    transition:
        background-color
        0.2s
        ease,
        color
        0.2s
        ease;
}


.tab button:hover {

    background:#E5E7EB;
}


.tab button.active {

    background:#3B4992;

    color:#FFFFFF;

    font-weight:600;
}


/* =========================================================
   搜索栏
   ========================================================= */

.species-search-panel {

    position:sticky;

    z-index:99998;

    width:100%;

    padding:
        9px
        18px
        9px
        18px;

    background:
        rgba(
            255,
            255,
            255,
            0.98
        );

    border-bottom:
        1px
        solid
        #E5E7EB;

    box-shadow:
        0
        2px
        7px
        rgba(
            0,
            0,
            0,
            0.035
        );
}


.species-search-main {

    position:relative;

    display:flex;

    align-items:center;

    width:min(
        920px,
        96%
    );

    margin:0 auto;

    border:
        1px
        solid
        #D1D5DB;

    border-radius:9px;

    background:#FFFFFF;

    transition:
        border-color
        .18s
        ease,
        box-shadow
        .18s
        ease;
}


.species-search-main:focus-within {

    border-color:#3B4992;

    box-shadow:
        0
        0
        0
        3px
        rgba(
            59,
            73,
            146,
            0.10
        );
}


.species-search-icon {

    flex:
        0
        0
        auto;

    padding-left:13px;

    font-size:16px;

    opacity:.70;
}


.species-search-input {

    flex:1;

    min-width:0;

    height:42px;

    padding:
        0
        12px;

    border:none;

    outline:none;

    background:
        transparent;

    font-size:14px;

    color:#1F2937;
}


.species-search-input::placeholder {

    color:#9CA3AF;
}


.species-search-button {

    flex:
        0
        0
        auto;

    height:34px;

    margin-right:4px;

    padding:
        0
        17px;

    border:none;

    border-radius:6px;

    background:#3B4992;

    color:#FFFFFF;

    font-weight:600;

    cursor:pointer;
}


.species-search-button:hover {

    opacity:.88;
}


.species-search-clear {

    flex:
        0
        0
        auto;

    height:34px;

    margin-right:4px;

    padding:
        0
        12px;

    border:none;

    border-radius:6px;

    background:#F3F4F6;

    color:#4B5563;

    cursor:pointer;
}


.species-search-clear:hover {

    background:#E5E7EB;
}


/* =========================================================
   搜索提示文字不再显示
   ========================================================= */

.species-search-status {

    display:none !important;
}


/* =========================================================
   搜索结果
   ========================================================= */

.species-search-results {

    display:none;

    position:absolute;

    left:50%;

    top:
        calc(
            100%
            -
            2px
        );

    transform:
        translateX(-50%);

    width:min(
        920px,
        calc(
            96%
            -
            36px
        )
    );

    max-height:
        min(
            430px,
            56vh
        );

    overflow-y:auto;

    margin:0;

    padding:6px;

    background:#FFFFFF;

    border:
        1px
        solid
        #E5E7EB;

    border-radius:9px;

    box-shadow:
        0
        10px
        30px
        rgba(
            0,
            0,
            0,
            .15
        );
}


.species-result-item {

    display:block;

    width:100%;

    padding:
        10px
        12px;

    border:none;

    border-bottom:
        1px
        solid
        #F1F3F5;

    border-radius:6px;

    background:#FFFFFF;

    text-align:left;

    cursor:pointer;
}


.species-result-item:last-child {

    border-bottom:none;
}


.species-result-item:hover {

    background:#F6F7FC;
}


.species-result-title {

    display:flex;

    align-items:center;

    flex-wrap:wrap;

    gap:8px;

    margin-bottom:4px;
}


.species-result-en {

    color:#3B4992;

    font-size:14px;

    font-weight:700;

    font-style:italic;
}


.species-result-cn {

    color:#222;

    font-size:13px;

    font-weight:600;
}


.species-result-source {

    display:inline-flex;

    align-items:center;

    padding:
        2px
        7px;

    border-radius:10px;

    background:#EEF2FF;

    color:#3B4992;

    font-size:10px;

    font-weight:600;
}


.species-result-category {

    display:inline-flex;

    align-items:center;

    padding:
        2px
        7px;

    border-radius:10px;

    background:#ECFDF5;

    color:#047857;

    font-size:10px;
}


.species-result-info {

    color:#6B7280;

    font-size:11px;

    line-height:1.6;
}


.species-result-id {

    margin-top:2px;

    color:#008B45;

    font-size:11px;
}


.species-search-empty {

    padding:
        25px
        15px;

    text-align:center;

    color:#9CA3AF;

    font-size:13px;
}


/* =========================================================
   ECharts
   ========================================================= */

.chart-container {

    width:100% !important;

    max-width:100vw !important;

    margin:0 auto;

    box-sizing:border-box;

    /*
       动画过程中，
       不允许极少量绘制像素跑到容器外。
    */
    overflow:hidden;

    /*
       独立浏览器绘制层，
       避免隐藏 Tab 与当前 Tab 发生视觉叠加。
    */
    isolation:isolate;
}


/* Canvas 不允许出现浏览器额外边框 */

.chart-container canvas {

    border:none !important;

    outline:none !important;
}


/* =========================================================
   韦恩图
   ========================================================= */

#__VENN_ID__ {

    width:100%;

    text-align:center;

    padding:
        15px
        20px
        30px
        20px;

    background:#FFFFFF;
}


.venn-action-bar {

    display:flex;

    align-items:center;

    justify-content:center;

    width:100%;

    margin:
        10px
        auto
        0
        auto;
}


.venn-download-btn {

    display:inline-flex;

    align-items:center;

    justify-content:center;

    padding:
        9px
        18px;

    border:
        1px
        solid
        #3B4992;

    border-radius:6px;

    background:#FFFFFF;

    color:#3B4992;

    font-size:14px;

    font-weight:600;

    cursor:pointer;
}


.venn-download-btn:hover {

    background:#3B4992;

    color:#FFFFFF;
}


/* =========================================================
   1000px 以下
   ========================================================= */

@media (
    max-width:1000px
) {

    .tab {

        justify-content:flex-start;

        flex-wrap:nowrap;

        overflow-x:auto;

        -webkit-overflow-scrolling:touch;
    }


    .tab button {

        flex:
            0
            0
            auto;

        min-width:155px;

        max-width:none;
    }

}


/* =========================================================
   手机
   ========================================================= */

@media (
    max-width:600px
) {

    .tab {

        padding:4px;
    }


    .tab button {

        min-width:145px;

        padding:
            8px
            10px;

        font-size:12px;
    }


    .species-search-panel {

        padding:6px;
    }


    .species-search-main {

        width:100%;
    }


    .species-search-results {

        width:
            calc(
                100%
                -
                12px
            );

        max-height:55vh;
    }


    .species-search-input {

        font-size:12px;

        padding:
            0
            7px;
    }


    .species-search-button {

        padding:
            0
            10px;

        font-size:12px;
    }


    .species-search-clear {

        padding:
            0
            8px;

        font-size:12px;
    }


    #__VENN_ID__ {

        padding:
            10px
            5px
            20px
            5px;
    }

}

</style>

"""


responsive_css = (
    responsive_css.replace(
        "__VENN_ID__",
        venn_chart_id
    )
)


html = html.replace(

    "</head>",

    responsive_css
    +
    "\n</head>",

    1
)


# ============================================================
# 21. 韦恩图下载按钮
# ============================================================

venn_button_html = """

<div class="venn-action-bar">

    <button
        class="venn-download-btn"
        type="button"
        onclick="downloadVennImage()"
    >
        ↓ 下载韦恩图 PNG
    </button>

</div>

"""


venn_pattern = (

    r'(<div[^>]*id="'
    +
    re.escape(
        venn_chart_id
    )
    +
    r'"[^>]*>)'
)


html = re.sub(

    venn_pattern,

    r'\1'
    +
    venn_button_html,

    html,

    count=1
)


# ============================================================
# 22. 响应式 + 边缘杂线优化 JavaScript
# ============================================================
#
# 核心原则：
#
# 1. 不关闭 Sunburst 原有动画。
#
# 2. Tab 切换时只 resize 当前显示的图。
#
# 3. resize 本身不使用尺寸渐变动画，
#    避免：
#    Sunburst 动画 + resize 动画
#    同时进行造成外圈杂线。
#
# 4. 不再连续多次无条件 resize。
#
# ============================================================

responsive_js = """

<script>

(function() {

    const chartIds =
        __CHART_IDS__;


    const vennChartId =
        "__VENN_ID__";


    let windowResizeTimer =
        null;


    let tabResizeTimer =
        null;


    // =====================================================
    // 获取布局信息
    // =====================================================

    function getLayoutInfo() {

        const tab =
            document.querySelector(
                ".tab"
            );


        const tabHeight =
            tab
            ?
            tab.offsetHeight
            :
            60;


        const searchPanel =
            document.getElementById(
                "species-search-panel"
            );


        const searchHeight =
            searchPanel
            ?
            searchPanel.offsetHeight
            :
            0;


        if (searchPanel) {

            searchPanel.style.top =
                tabHeight
                +
                "px";
        }


        const availableHeight =
            Math.max(

                420,

                window.innerHeight
                -
                tabHeight
                -
                searchHeight
                -
                12
            );


        return {

            tabHeight:
                tabHeight,

            searchHeight:
                searchHeight,

            availableHeight:
                availableHeight
        };
    }


    // =====================================================
    // 设置图表 DOM 尺寸
    // =====================================================

    function updateContainerSizes() {

        const layout =
            getLayoutInfo();


        chartIds.forEach(

            function(id) {

                const dom =
                    document.getElementById(
                        id
                    );


                if (!dom) {
                    return;
                }


                dom.style.width =
                    "100%";


                dom.style.height =
                    layout.availableHeight
                    +
                    "px";
            }
        );


        const vennContainer =
            document.getElementById(
                vennChartId
            );


        if (vennContainer) {

            vennContainer.style.minHeight =
                layout.availableHeight
                +
                "px";
        }


        return (
            layout.availableHeight
        );
    }


    // =====================================================
    // Resize 当前可见图
    // =====================================================

    function resizeVisibleChart(
        force
    ) {

        updateContainerSizes();


        chartIds.forEach(

            function(id) {

                const dom =
                    document.getElementById(
                        id
                    );


                /*
                 * 隐藏标签页绝对不 resize。
                 */
                if (
                    !dom
                    ||
                    dom.offsetParent === null
                ) {

                    return;
                }


                try {

                    const chart =
                        echarts.getInstanceByDom(
                            dom
                        );


                    if (!chart) {
                        return;
                    }


                    const targetWidth =
                        Math.max(
                            1,
                            dom.clientWidth
                        );


                    const targetHeight =
                        Math.max(
                            1,
                            dom.clientHeight
                        );


                    const currentWidth =
                        chart.getWidth();


                    const currentHeight =
                        chart.getHeight();


                    const widthChanged =
                        Math.abs(
                            currentWidth
                            -
                            targetWidth
                        )
                        >
                        2;


                    const heightChanged =
                        Math.abs(
                            currentHeight
                            -
                            targetHeight
                        )
                        >
                        2;


                    if (
                        force
                        ||
                        widthChanged
                        ||
                        heightChanged
                    ) {

                        /*
                         * 重要：
                         *
                         * 这里只取消“resize 尺寸变化动画”。
                         *
                         * 没有关闭 Sunburst 自身动画：
                         *
                         * - 初始动画
                         * - 点击钻取动画
                         * - rootToNode 动画
                         * - highlight 动画
                         *
                         * 全部保留。
                         */

                        chart.resize({

                            width:
                                targetWidth,

                            height:
                                targetHeight,

                            silent:
                                true,

                            animation: {

                                duration:
                                    0
                            }
                        });
                    }

                }

                catch(error) {

                    console.warn(
                        "Chart resize error:",
                        error
                    );
                }

            }
        );
    }


    // =====================================================
    // 对外提供布局调整
    // =====================================================

    function resizeAllCharts() {

        updateContainerSizes();


        requestAnimationFrame(

            function() {

                requestAnimationFrame(

                    function() {

                        resizeVisibleChart(
                            false
                        );
                    }
                );
            }
        );
    }


    window.resizeAllCharts =
        resizeAllCharts;


    // =====================================================
    // 下载韦恩图
    // =====================================================

    window.downloadVennImage =
        function() {

            const container =
                document.getElementById(
                    vennChartId
                );


            if (!container) {

                alert(
                    "未找到韦恩图容器。"
                );

                return;
            }


            const image =
                container.querySelector(
                    "img"
                );


            if (!image) {

                alert(
                    "未找到韦恩图图片。"
                );

                return;
            }


            const link =
                document.createElement(
                    "a"
                );


            link.href =
                image.src;


            link.download =
                "07_Venn_Adult_vs_Child.png";


            document.body.appendChild(
                link
            );


            link.click();


            document.body.removeChild(
                link
            );
        };


    // =====================================================
    // 首次页面加载
    // =====================================================

    window.addEventListener(

        "load",

        function() {

            const tabs =
                document.querySelectorAll(
                    ".tablinks"
                );


            if (
                tabs
                &&
                tabs.length > 0
            ) {

                tabs[0].classList.add(
                    "active"
                );
            }


            /*
             * 等浏览器真正完成页面布局后
             * 再校正当前显示图。
             */

            requestAnimationFrame(

                function() {

                    requestAnimationFrame(

                        function() {

                            resizeVisibleChart(
                                true
                            );
                        }
                    );
                }
            );
        }
    );


    // =====================================================
    // 浏览器窗口尺寸变化
    // =====================================================

    window.addEventListener(

        "resize",

        function() {

            clearTimeout(
                windowResizeTimer
            );


            windowResizeTimer =
                setTimeout(

                    function() {

                        resizeVisibleChart(
                            false
                        );

                    },

                    120
                );
        }
    );


    // =====================================================
    // Tab 切换
    // =====================================================

    document.addEventListener(

        "click",

        function(event) {

            const target =
                event.target.closest(
                    ".tablinks"
                );


            if (!target) {
                return;
            }


            clearTimeout(
                tabResizeTimer
            );


            /*
             * 先更新 DOM 尺寸，
             * 不立即操作 Canvas。
             */

            updateContainerSizes();


            /*
             * 等 Tab display 状态稳定两个 frame。
             */

            requestAnimationFrame(

                function() {

                    requestAnimationFrame(

                        function() {

                            /*
                             * 当前图只校正一次。
                             *
                             * resize 不做过渡动画，
                             * Sunburst 自身动画仍保留。
                             */

                            resizeVisibleChart(
                                true
                            );


                            /*
                             * 一段时间以后只做尺寸检查。
                             *
                             * 如果尺寸没变化，
                             * 不会再次调用 chart.resize()。
                             */

                            tabResizeTimer =
                                setTimeout(

                                    function() {

                                        resizeVisibleChart(
                                            false
                                        );

                                    },

                                    900
                                );
                        }
                    );
                }
            );

        }
    );

})();

</script>

"""


responsive_js = responsive_js.replace(

    "__CHART_IDS__",

    chart_ids_json

).replace(

    "__VENN_ID__",

    venn_chart_id
)


html = html.replace(

    "</body>",

    responsive_js
    +
    "\n</body>",

    1
)


# ============================================================
# 23. 菌种搜索 JavaScript
# ============================================================

search_js = """

<script>

(function() {

    const speciesDatabase =
        __SEARCH_DATA__;


    const chartTargets =
        __CHART_TARGETS__;


    let currentSearchResults =
        [];


    // =====================================================
    // HTML 转义
    // =====================================================

    function escapeHtml(value) {

        if (
            value === null
            ||
            value === undefined
        ) {

            return "";
        }


        return String(value)

            .replace(
                /&/g,
                "&amp;"
            )

            .replace(
                /</g,
                "&lt;"
            )

            .replace(
                />/g,
                "&gt;"
            )

            .replace(
                /"/g,
                "&quot;"
            )

            .replace(
                /'/g,
                "&#039;"
            );
    }


    // =====================================================
    // 标准化搜索文字
    // =====================================================

    function normalizeText(value) {

        return String(
            value || ""
        )
        .trim()
        .toLowerCase();
    }


    // =====================================================
    // 完全匹配
    // =====================================================

    function arrayExact(
        values,
        query
    ) {

        return (
            values || []
        ).some(

            function(value) {

                return (
                    normalizeText(
                        value
                    )
                    ===
                    query
                );
            }
        );
    }


    // =====================================================
    // 前缀匹配
    // =====================================================

    function arrayStartsWith(
        values,
        query
    ) {

        return (
            values || []
        ).some(

            function(value) {

                return normalizeText(
                    value
                ).startsWith(
                    query
                );
            }
        );
    }


    // =====================================================
    // 包含匹配
    // =====================================================

    function arrayIncludes(
        values,
        query
    ) {

        return (
            values || []
        ).some(

            function(value) {

                return normalizeText(
                    value
                ).includes(
                    query
                );
            }
        );
    }


    // =====================================================
    // 搜索评分
    //
    // 0 = 完全匹配
    // 1 = 前缀匹配
    // 2 = 包含匹配
    // =====================================================

    function getSearchScore(
        item,
        query
    ) {

        const q =
            normalizeText(
                query
            );


        if (!q) {

            return 999;
        }


        // -------------------------------------------------
        // 完全匹配
        // -------------------------------------------------

        if (
            arrayExact(
                item.strain_ids,
                q
            )
            ||
            arrayExact(
                item.chinese_aliases,
                q
            )
            ||
            arrayExact(
                item.english_aliases,
                q
            )
            ||
            normalizeText(
                item.merge_key
            )
            ===
            q
        ) {

            return 0;
        }


        // -------------------------------------------------
        // 前缀匹配
        // -------------------------------------------------

        if (
            arrayStartsWith(
                item.strain_ids,
                q
            )
            ||
            arrayStartsWith(
                item.chinese_aliases,
                q
            )
            ||
            arrayStartsWith(
                item.english_aliases,
                q
            )
            ||
            normalizeText(
                item.merge_key
            ).startsWith(
                q
            )
        ) {

            return 1;
        }


        // -------------------------------------------------
        // 包含匹配
        // -------------------------------------------------

        if (
            arrayIncludes(
                item.strain_ids,
                q
            )
            ||
            arrayIncludes(
                item.chinese_aliases,
                q
            )
            ||
            arrayIncludes(
                item.english_aliases,
                q
            )
            ||
            normalizeText(
                item.merge_key
            ).includes(
                q
            )
        ) {

            return 2;
        }


        return 999;
    }


    // =====================================================
    // 执行搜索
    // =====================================================

    function searchSpecies(
        query
    ) {

        const result =
            [];


        speciesDatabase.forEach(

            function(item) {

                const score =
                    getSearchScore(
                        item,
                        query
                    );


                if (
                    score
                    <
                    999
                ) {

                    result.push({

                        item:
                            item,

                        score:
                            score
                    });
                }
            }
        );


        result.sort(

            function(a, b) {

                if (
                    a.score
                    !==
                    b.score
                ) {

                    return (
                        a.score
                        -
                        b.score
                    );
                }


                return String(
                    a.item.english_name
                    ||
                    ""
                ).localeCompare(

                    String(
                        b.item.english_name
                        ||
                        ""
                    )
                );
            }
        );


        return result.map(

            function(resultItem) {

                return (
                    resultItem.item
                );
            }
        );
    }


    // =====================================================
    // 韦恩分类名称
    // =====================================================

    function getCategoryText(
        item
    ) {

        if (
            item.venn_class
            ===
            "成人特有 (Adult Only)"
        ) {

            return "成人特有";
        }


        if (
            item.venn_class
            ===
            "儿童特有 (Child Only)"
        ) {

            return "儿童特有";
        }


        if (
            item.venn_class
            ===
            "交集 (Intersection)"
        ) {

            return "成人儿童共有";
        }


        return "";
    }


    // =====================================================
    // 搜索结果
    // =====================================================

    function renderSearchResults(
        results
    ) {

        const resultsBox =
            document.getElementById(
                "species-search-results"
            );


        const status =
            document.getElementById(
                "species-search-status"
            );


        if (!resultsBox) {
            return;
        }


        currentSearchResults =
            results;


        if (
            results.length
            ===
            0
        ) {

            resultsBox.innerHTML = `

                <div
                    class="species-search-empty"
                >
                    未找到匹配菌种
                </div>

            `;


            resultsBox.style.display =
                "block";


            if (status) {

                status.textContent =
                    "没有找到匹配结果";
            }


            return;
        }


        if (status) {

            status.textContent =
                "找到 "
                +
                results.length
                +
                " 个匹配菌种";
        }


        const displayResults =
            results.slice(
                0,
                30
            );


        let html =
            "";


        displayResults.forEach(

            function(item, index) {

                const strainText =
                    (
                        item.strain_ids
                        &&
                        item.strain_ids.length
                    )
                    ?
                    item.strain_ids.join(
                        ", "
                    )
                    :
                    "无";


                const categoryText =
                    getCategoryText(
                        item
                    );


                html += `

                <button
                    type="button"
                    class="species-result-item"
                    data-result-index="${index}"
                >

                    <div
                        class="species-result-title"
                    >

                        <span
                            class="species-result-en"
                        >
                            ${
                                escapeHtml(
                                    item.english_name
                                )
                            }
                        </span>

                        <span
                            class="species-result-cn"
                        >
                            ${
                                escapeHtml(
                                    item.chinese_name
                                )
                            }
                        </span>

                        <span
                            class="species-result-source"
                        >
                            ${
                                escapeHtml(
                                    item.source
                                )
                            }
                        </span>

                        ${
                            categoryText
                            ?
                            `
                            <span
                                class="species-result-category"
                            >
                                ${
                                    escapeHtml(
                                        categoryText
                                    )
                                }
                            </span>
                            `
                            :
                            ""
                        }

                    </div>

                    <div
                        class="species-result-info"
                    >

                        ${
                            escapeHtml(
                                item.phylum
                            )
                        }

                        &nbsp;›&nbsp;

                        ${
                            escapeHtml(
                                item.family
                            )
                        }

                        &nbsp;›&nbsp;

                        ${
                            escapeHtml(
                                item.genus
                            )
                        }

                    </div>

                    <div
                        class="species-result-id"
                    >

                        新菌序号：

                        ${
                            escapeHtml(
                                strainText
                            )
                        }

                    </div>

                </button>

                `;
            }
        );


        resultsBox.innerHTML =
            html;


        resultsBox.style.display =
            "block";
    }


    // =====================================================
    // 判断菌种应该进入哪个图
    // =====================================================

    function getTargetChartId(
        item
    ) {

        if (
            item.venn_class
            ===
            "成人特有 (Adult Only)"
        ) {

            return (
                chartTargets.adult_only
            );
        }


        if (
            item.venn_class
            ===
            "儿童特有 (Child Only)"
        ) {

            return (
                chartTargets.child_only
            );
        }


        if (
            item.venn_class
            ===
            "交集 (Intersection)"
        ) {

            return (
                chartTargets.intersection
            );
        }


        return (
            chartTargets.overview
        );
    }


    // =====================================================
    // 根据 Chart ID 切换 Tab
    // =====================================================

    function switchToChartTab(
        chartId
    ) {

        const buttons =
            document.querySelectorAll(
                ".tablinks"
            );


        for (
            let i = 0;
            i < buttons.length;
            i++
        ) {

            const onclickText =
                buttons[i].getAttribute(
                    "onclick"
                )
                ||
                "";


            if (
                onclickText.includes(
                    chartId
                )
            ) {

                buttons[i].click();

                return true;
            }
        }


        return false;
    }


    // =====================================================
    // 在 ECharts Tree 中寻找菌种
    // =====================================================

    function findSpeciesNode(
        chart,
        searchId
    ) {

        if (!chart) {
            return null;
        }


        const model =
            chart.getModel();


        const seriesModel =
            model.getSeriesByIndex(
                0
            );


        if (!seriesModel) {
            return null;
        }


        const data =
            seriesModel.getData();


        if (
            !data
            ||
            !data.tree
            ||
            !data.tree.root
        ) {

            return null;
        }


        let targetNode =
            null;


        data.tree.root.eachNode(

            function(node) {

                if (
                    targetNode
                    ||
                    node.dataIndex < 0
                ) {

                    return;
                }


                let rawItem =
                    null;


                try {

                    rawItem =
                        data.getRawDataItem(
                            node.dataIndex
                        );

                }

                catch(error) {

                    rawItem =
                        null;
                }


                if (
                    rawItem
                    &&
                    rawItem.search_id
                    ===
                    searchId
                ) {

                    targetNode =
                        node;
                }
            }
        );


        if (!targetNode) {
            return null;
        }


        return {

            node:
                targetNode,

            data:
                data,

            seriesModel:
                seriesModel
        };
    }


    // =====================================================
    // 定位菌种
    // =====================================================

    function locateSpecies(
        item
    ) {

        const resultsBox =
            document.getElementById(
                "species-search-results"
            );


        if (resultsBox) {

            resultsBox.style.display =
                "none";
        }


        const chartId =
            getTargetChartId(
                item
            );


        /*
         * switchToChartTab 本身会触发
         * responsive_js 中的 Tab 尺寸校正。
         *
         * 所以这里不再额外立即 resize，
         * 避免重复 Canvas 重绘。
         */

        switchToChartTab(
            chartId
        );


        setTimeout(

            function() {

                const dom =
                    document.getElementById(
                        chartId
                    );


                if (!dom) {

                    alert(
                        "没有找到对应旭日图。"
                    );

                    return;
                }


                const chart =
                    echarts.getInstanceByDom(
                        dom
                    );


                if (!chart) {

                    alert(
                        "图表尚未完成初始化。"
                    );

                    return;
                }


                const found =
                    findSpeciesNode(
                        chart,
                        item.search_id
                    );


                if (!found) {

                    if (
                        chartId
                        !==
                        chartTargets.overview
                    ) {

                        locateSpeciesInOverview(
                            item
                        );

                        return;
                    }


                    alert(
                        "搜索到了该菌种，但旭日图中没有找到对应节点。"
                    );

                    return;
                }


                performNodeLocation(
                    chart,
                    found,
                    item
                );

            },

            320
        );
    }


    // =====================================================
    // 回退到总体图
    // =====================================================

    function locateSpeciesInOverview(
        item
    ) {

        const chartId =
            chartTargets.overview;


        switchToChartTab(
            chartId
        );


        setTimeout(

            function() {

                const dom =
                    document.getElementById(
                        chartId
                    );


                if (!dom) {
                    return;
                }


                const chart =
                    echarts.getInstanceByDom(
                        dom
                    );


                if (!chart) {
                    return;
                }


                const found =
                    findSpeciesNode(
                        chart,
                        item.search_id
                    );


                if (found) {

                    performNodeLocation(
                        chart,
                        found,
                        item
                    );
                }

            },

            320
        );
    }


    // =====================================================
    // 钻取动画完成以后执行
    //
    // 不再死写 550ms / 760ms。
    //
    // 优先监听 ECharts finished，
    // 可以适应不同浏览器和不同动画速度。
    // =====================================================

    function runAfterChartFinished(
        chart,
        callback
    ) {

        let completed =
            false;


        let fallbackTimer =
            null;


        function finishOnce() {

            if (completed) {
                return;
            }


            completed =
                true;


            if (fallbackTimer) {

                clearTimeout(
                    fallbackTimer
                );
            }


            try {

                chart.off(
                    "finished",
                    finishOnce
                );

            }

            catch(error) {

                // 忽略
            }


            callback();
        }


        try {

            chart.on(
                "finished",
                finishOnce
            );

        }

        catch(error) {

            // 如果当前 ECharts 环境不支持，
            // 下面 fallback 仍然有效。
        }


        /*
         * 极端情况下 finished 没有触发，
         * 1.4 秒后自动执行。
         */

        fallbackTimer =
            setTimeout(
                finishOnce,
                1400
            );
    }


    // =====================================================
    // 真正执行定位
    //
    // 1. 钻取到属
    // 2. 保留原有动画
    // 3. 动画结束后高亮 species
    // 4. 显示 Tooltip
    // =====================================================

    function performNodeLocation(
        chart,
        found,
        item
    ) {

        const speciesNode =
            found.node;


        let genusNode =
            speciesNode.parentNode;


        if (!genusNode) {

            genusNode =
                speciesNode;
        }


        // -------------------------------------------------
        // 清除之前的高亮
        // -------------------------------------------------

        try {

            chart.dispatchAction({

                type:
                    "downplay",

                seriesIndex:
                    0
            });

        }

        catch(error) {

            console.warn(
                "Downplay error:",
                error
            );
        }


        // -------------------------------------------------
        // 动画结束后的动作
        // -------------------------------------------------

        runAfterChartFinished(

            chart,

            function() {

                const dataIndex =
                    speciesNode.dataIndex;


                // -----------------------------------------
                // 高亮
                // -----------------------------------------

                try {

                    chart.dispatchAction({

                        type:
                            "highlight",

                        seriesIndex:
                            0,

                        dataIndex:
                            dataIndex
                    });

                }

                catch(error) {

                    console.warn(
                        "Highlight error:",
                        error
                    );
                }


                // -----------------------------------------
                // Tooltip
                // -----------------------------------------

                try {

                    chart.dispatchAction({

                        type:
                            "showTip",

                        seriesIndex:
                            0,

                        dataIndex:
                            dataIndex
                    });

                }

                catch(error) {

                    console.warn(
                        "Tooltip error:",
                        error
                    );
                }


                const status =
                    document.getElementById(
                        "species-search-status"
                    );


                if (status) {

                    const category =
                        getCategoryText(
                            item
                        );


                    status.innerHTML =

                        "已定位："

                        +

                        "<b style='" +
                        "color:#3B4992;" +
                        "font-style:italic;" +
                        "'>"

                        +

                        escapeHtml(
                            item.english_name
                        )

                        +

                        "</b>"

                        +

                        " ｜ "

                        +

                        escapeHtml(
                            item.chinese_name
                        )

                        +

                        (
                            category
                            ?
                            " ｜ "
                            +
                            escapeHtml(
                                category
                            )
                            :
                            ""
                        );
                }
            }
        );


        // -------------------------------------------------
        // 钻取到“属”
        //
        // 这里没有关闭 animation。
        // -------------------------------------------------

        try {

            chart.dispatchAction({

                type:
                    "sunburstRootToNode",

                seriesIndex:
                    0,

                targetNode:
                    genusNode
            });

        }

        catch(error) {

            console.warn(
                "Sunburst drill error:",
                error
            );
        }
    }


    // =====================================================
    // 搜索
    // =====================================================

    function executeSearch(
        autoLocate
    ) {

        const input =
            document.getElementById(
                "species-search-input"
            );


        const resultsBox =
            document.getElementById(
                "species-search-results"
            );


        const status =
            document.getElementById(
                "species-search-status"
            );


        if (!input) {
            return;
        }


        const query =
            input.value.trim();


        if (!query) {

            currentSearchResults =
                [];


            if (resultsBox) {

                resultsBox.style.display =
                    "none";
            }


            if (status) {

                status.textContent =
                    "";
            }


            return;
        }


        const results =
            searchSpecies(
                query
            );


        renderSearchResults(
            results
        );


        if (!autoLocate) {
            return;
        }


        // -------------------------------------------------
        // 唯一结果直接定位
        // -------------------------------------------------

        if (
            results.length
            ===
            1
        ) {

            locateSpecies(
                results[0]
            );

            return;
        }


        // -------------------------------------------------
        // 多结果检查唯一完全匹配
        // -------------------------------------------------

        const q =
            normalizeText(
                query
            );


        const exactResults =
            results.filter(

                function(item) {

                    return (

                        arrayExact(
                            item.strain_ids,
                            q
                        )

                        ||

                        arrayExact(
                            item.chinese_aliases,
                            q
                        )

                        ||

                        arrayExact(
                            item.english_aliases,
                            q
                        )

                        ||

                        normalizeText(
                            item.merge_key
                        )
                        ===
                        q
                    );
                }
            );


        if (
            exactResults.length
            ===
            1
        ) {

            locateSpecies(
                exactResults[0]
            );
        }
    }


    // =====================================================
    // 页面事件
    // =====================================================

    window.addEventListener(

        "load",

        function() {

            const panel =
                document.getElementById(
                    "species-search-panel"
                );


            const input =
                document.getElementById(
                    "species-search-input"
                );


            const searchButton =
                document.getElementById(
                    "species-search-button"
                );


            const clearButton =
                document.getElementById(
                    "species-search-clear"
                );


            const resultsBox =
                document.getElementById(
                    "species-search-results"
                );


            if (!input) {
                return;
            }


            // -------------------------------------------------
            // 实时搜索
            // -------------------------------------------------

            input.addEventListener(

                "input",

                function() {

                    executeSearch(
                        false
                    );
                }
            );


            // -------------------------------------------------
            // Enter / Esc
            // -------------------------------------------------

            input.addEventListener(

                "keydown",

                function(event) {

                    if (
                        event.key
                        ===
                        "Enter"
                    ) {

                        event.preventDefault();

                        executeSearch(
                            true
                        );
                    }


                    if (
                        event.key
                        ===
                        "Escape"
                    ) {

                        if (resultsBox) {

                            resultsBox.style.display =
                                "none";
                        }
                    }
                }
            );


            // -------------------------------------------------
            // 搜索按钮
            // -------------------------------------------------

            if (searchButton) {

                searchButton.addEventListener(

                    "click",

                    function() {

                        executeSearch(
                            true
                        );
                    }
                );
            }


            // -------------------------------------------------
            // 清空按钮
            // -------------------------------------------------

            if (clearButton) {

                clearButton.addEventListener(

                    "click",

                    function() {

                        input.value =
                            "";


                        currentSearchResults =
                            [];


                        if (resultsBox) {

                            resultsBox.style.display =
                                "none";
                        }


                        const status =
                            document.getElementById(
                                "species-search-status"
                            );


                        if (status) {

                            status.textContent =
                                "";
                        }


                        input.focus();
                    }
                );
            }


            // -------------------------------------------------
            // 点击搜索结果
            // -------------------------------------------------

            if (resultsBox) {

                resultsBox.addEventListener(

                    "click",

                    function(event) {

                        const button =
                            event.target.closest(
                                ".species-result-item"
                            );


                        if (!button) {
                            return;
                        }


                        const index =
                            Number(
                                button.dataset.resultIndex
                            );


                        const item =
                            currentSearchResults[
                                index
                            ];


                        if (item) {

                            locateSpecies(
                                item
                            );
                        }
                    }
                );
            }


            // -------------------------------------------------
            // 点击外部关闭搜索结果
            // -------------------------------------------------

            document.addEventListener(

                "click",

                function(event) {

                    if (
                        panel
                        &&
                        !panel.contains(
                            event.target
                        )
                    ) {

                        if (resultsBox) {

                            resultsBox.style.display =
                                "none";
                        }
                    }
                }
            );


            // -------------------------------------------------
            // 再次点击输入框恢复搜索结果
            // -------------------------------------------------

            input.addEventListener(

                "focus",

                function() {

                    if (
                        currentSearchResults.length
                        >
                        0
                        &&
                        resultsBox
                    ) {

                        resultsBox.style.display =
                            "block";
                    }
                }
            );

        }
    );

})();

</script>

"""


search_js = search_js.replace(

    "__SEARCH_DATA__",

    search_data_json

).replace(

    "__CHART_TARGETS__",

    chart_targets_json
)


html = html.replace(

    "</body>",

    search_js
    +
    "\n</body>",

    1
)


# ============================================================
# 24. 写回最终 HTML
# ============================================================

with open(
    output_file,
    'w',
    encoding='utf-8'
) as f:

    f.write(
        html
    )


# ============================================================
# 25. 完成提示
# ============================================================

print("")
print("========================================")
print("生成成功！")
print("----------------------------------------")
print(f"HTML：{output_file}")
print("----------------------------------------")

print(
    f"① 总体分布 ({total_species_count}种)"
)

print(
    f"② 全部成人 ({adult_species_count}种)"
)

print(
    f"③ 全部儿童 ({child_species_count}种)"
)

print(
    f"④ 仅成人 ({adult_only_count}种)"
)

print(
    f"⑤ 仅儿童 ({child_only_count}种)"
)

print(
    f"⑥ 共有交集 ({intersection_count}种)"
)

print(
    f"⑦ 韦恩图 ({total_species_count}种)"
)

print("----------------------------------------")
print("功能：")
print("✓ 浏览器窗口自适应")
print("✓ 旭日图高清 PNG 下载")
print("✓ 旭日图一键还原")
print("✓ 韦恩图嵌入 HTML")
print("✓ 韦恩图 PNG 下载")
print("✓ 中文菌名搜索")
print("✓ 英文菌名搜索")
print("✓ 新菌序号搜索")
print("✓ 实时模糊匹配")
print("✓ 搜索后自动切换分类页面")
print("✓ 自动钻取至菌种所在属")
print("✓ 自动高亮目标菌种")
print("✓ 自动显示目标菌种 Tooltip")
print("✓ 已去掉搜索框下方提示文字")
print("✓ 已优化 Tab 切换边缘杂乱线条")
print("✓ 已减少重复 resize")
print("✓ resize 不再与旭日图动画叠加")
print("✓ 保留旭日图原有动画")
print("✓ 保留点击钻取动画")
print("✓ 保留搜索自动钻取动画")
print("✓ 搜索高亮等待钻取动画结束")
print("✓ 不再使用 set_global_opts(animation_opts=...)")
print("========================================")
