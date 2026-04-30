# StarNet 配表说明

推荐维护 1 个 Excel 文件（两页签）：

- `starnet.xlsx`：
  - 页签 `profiles`：编号与主页映射 + 基础信息（简介、验证问题、答案）
  - 页签 `feed`：动态与评论（合并在一张表）

也兼容 CSV 兜底模式（无 `starnet.xlsx` 时读取 `profiles.csv` + `feed.csv`）。

## 字段说明

### profiles.csv
- `profile_id`: 主页编号（如 `P001`）
- `profile_type`: 主页类型（推荐填 `艺人` / `粉丝` / `黑粉`）
- `file`: 主页文件路径（相对仓库根目录，可留空让脚本自动生成）
- `profile_slug`: 文件 slug（如 `orbit_escape`，当 `file` 为空时用于生成文件名）
- `display_name`: 显示昵称（会写入页面 `h2.name`，并用于动态头部）
- `bio`: 简介
- `follow_question`: 关注验证问题（不需要写“问题：”前缀）
- `follow_answer`: 验证答案

## 新建主页（通过表）

当 `profiles` 里新增一行且对应文件不存在时，脚本会自动创建页面：

- 按 `profile_type` 自动放目录：
  - `艺人` -> `social/artists/`
  - `粉丝` -> `social/civilians/fans/`
  - `黑粉` -> `social/civilians/haters/`
- 若 `file` 留空，会按 `profile_slug` 生成：`starnet-social-{profile_slug}.html`
- 若 `profile_slug` 也为空，会用 `profile_id` 自动生成文件名

### feed.csv
- `profile_id`: 对应 `profiles.csv` 的编号
- `post_order`: 动态顺序（数字越小越靠前，通常 1=最新）
- `time`: 时间文案（如“昨天”“3 天前”）
- `text`: 动态正文
- `followers_only`: 填 `1` 表示“关注后可见”；留空表示公开可见
- `image`: 可选，配图路径（建议写仓库相对路径，如 `assets/xxx.png`）
- `image_alt`: 可选，图片说明（不填默认“动态配图”）
- `comment_nickname_1` / `comment_text_1` ... `comment_nickname_8` / `comment_text_8`:
  - 只填 `comment_text_x` 就会生成评论；
  - `comment_nickname_x` 留空会自动生成 `StarNetxxxx`；
  - `comment_text_x` 留空就视为没有该条评论。

## 自动生成互动数

你不需要填写转发/点赞/评论数。脚本会按动态类型自动生成：

- 关注后可见动态：转发偏低、赞评偏高（更像圈内帖）
- 公开动态：整体互动更高

## 执行更新

批量改 `profile_slug` 并同步文件名、站内链接、`starnet-home` 艺人入口：

`python3 scripts/apply_starnet_profile_slugs.py`

在仓库根目录运行（按表里内容写主页）：

`python3 scripts/update_starnet_profiles.py`

脚本会读取 `starnet.xlsx`（若存在），按表内容直接回写对应 HTML 页面。
