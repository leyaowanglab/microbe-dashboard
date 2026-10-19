# 菌种分布图（Microbe Dashboard）

基于鼻拭子微生物菌种数据（成人 / 儿童）生成交互式旭日图 + 韦恩图，并自动发布为在线网页。

## 在线访问

- **GitHub Pages 地址**：`https://leyaowanglab.smart.org.cn/microbe-dashboard/`
  （部署成功后，打开此地址即可看到最新网页）

## 项目文件说明

| 文件 | 作用 |
|------|------|
| `细菌库.xlsx` | 菌种原始数据（成人 / 儿童两个 sheet） |
| `generate_dashboard.py` | 核心脚本：读取数据 → 生成旭日图 + 韦恩图 HTML |
| `requirements.txt` | Python 依赖清单（GitHub Actions 用它安装环境） |
| `.github/workflows/build.yml` | GitHub Actions 自动构建 + 发布配置 |
| `watcher.py` | （可选）本地文件监听器，检测 细菌库.xlsx 变化自动重跑 |
| `启动监视器.bat` | （可选）双击启动本地监视器 |
| `手动重新生成.bat` | （可选）双击立即重新生成 HTML |

## 自动构建 + 发布流程

```
修改细菌库.xlsx
      ↓
git 提交并推送到 GitHub main 分支
      ↓
GitHub Actions 自动触发
      ↓
安装依赖 → 运行 generate_dashboard.py → 生成 HTML
      ↓
自动发布到 GitHub Pages
      ↓
在线网址即可看到最新网页
```

## 首次部署步骤（只需做一次）

### 1. 初始化并提交代码

```bash
# 进入项目目录
cd microbe_dashboard

# 初始化 git 仓库
git init

# 添加所有文件（.gitignore 会自动排除不需要的产物）
git add .

# 提交
git commit -m "初始提交：菌种分布图生成脚本"
```

### 2. 在 GitHub 新建仓库并推送

1. 打开 https://github.com/new 新建一个**公开**仓库（例如 `microbe-dashboard`）
   - 注意：GitHub Pages 免费版要求仓库为 public（私有仓库需付费）
2. 按照页面提示把本地代码推送上去：

```bash
git remote add origin https://github.com/<你的用户名>/<仓库名>.git
git branch -M main
git push -u origin main
```

### 3. 开启 GitHub Pages

1. 进入仓库 **Settings → Pages**
2. 在 **Build and deployment** 的 **Source** 下拉框中选择 **GitHub Actions**
3. 保存即可

第一次推送后，Actions 会自动运行，首次构建完成后 Pages 即生效。

## 日常更新数据（每次改数据都这样做）

```bash
# 修改好 细菌库.xlsx 后，执行：
git add 细菌库.xlsx
git commit -m "更新菌种数据"
git push
```

推送后 GitHub 会自动重新构建并更新网页，**无需任何其他操作**。

## 依赖说明

本地运行需要 Python 3.8+，并安装依赖：

```bash
pip install -r requirements.txt
python generate_dashboard.py
```

生成结果为：
- `菌种分布图（YYYYMMDD）.html` —— 交互式网页
