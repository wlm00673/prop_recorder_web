# 螺旋桨测试数据记录器（网页版）

纯静态网页（HTML + CSS + JS），无需服务器、无需构建。
原理与 `stol_calculator` 一样：**传到 GitHub Pages 就能给任何人打开**。

两种数据模式：

| 模式 | 怎么开启 | 效果 |
|---|---|---|
| **本机保存**（默认） | 什么都不用做 | 数据存在各自浏览器里，一人一份，页面上显示「本机保存」 |
| **云端共享** | 在 `config.js` 里填 Supabase 两项配置 | **所有人打开网页共享同一份数据**，页面上显示「云端共享 · N 条」 |

## 功能

- **桨尺寸**：固定翼 53 个 / 多旋翼 24 个常用尺寸一键选择，也可手动改直径、螺距
- **桨属性**：用途（固定翼 / 多旋翼）、动力（电动 / 油动）、材质（尼龙玻纤 / 碳纤维 / 榉木 / 尼龙 / 其他）
- **电机库**：内置朗宇(SunnySky) 14 款、T-Motor 16 款、新西达、银燕等共 36 个电机；可从库中选择，也可手动输入
- **电调库**：内置 T-Motor 66A 等 13 个电调；同样支持手动输入
- **自动记住**：手动输入过的电机 / 电调会自动加入下拉库；也可在页面里自己添加 / 删除
- **录入即保存**：点「保存记录」，或在「拉力」框里按回车（可开关），也支持 `Ctrl + S`
- **记录管理**：关键词 / 用途 / 材质 / 电机筛选，按电机汇总统计，删除单条或清空
- **导出**：Excel（.xlsx，含「测试记录」+「按电机汇总」两个工作表，纯前端生成）/ CSV / JSON 备份，可再导入 JSON 恢复
- **云端共享**：60 秒自动同步，也可点「刷新」立即拉取；连接失败会显示红色横幅提示
- **备注**：每条记录可写自定义备注（含电流、转速、电压、油门、温度、测试者、地点）

已预置：电机 **朗宇 X2820 III 1050KV（CUADC 专用）**、电调 **T-Motor 66A ESC**。

## 文件

| 文件 | 说明 |
|---|---|
| `index.html` | 页面结构 |
| `style.css` | 样式 |
| `app.js` | 全部逻辑（存储、记录、筛选、Excel 导出、云端同步） |
| `config.js` | **配置文件**：填 Supabase 的 URL / Key 即开启云端共享（不填=本机保存） |
| `README.md` | 本说明 |

## 本地预览

直接双击 `index.html`。若浏览器因 `file://` 限制禁用本地存储，用本地服务器：

```bash
cd prop_record_web
python3 -m http.server 8080
# 浏览器打开 http://localhost:8080
```

## 部署到 GitHub Pages

1. GitHub 新建仓库（Public），例如 `prop_recorder_web`
2. 把 `index.html`、`style.css`、`app.js`、`config.js`（`README.md` 可选）上传到仓库**根目录**
3. 仓库 → **Settings → Pages** → Source 选 **Deploy from a branch** → Branch 选 `main` + `/ (root)` → **Save**
4. 等 1~2 分钟，访问：`https://你的用户名.github.io/prop_recorder_web/`

---

## 让所有人共享同一份数据（Supabase，免费）

### 第 1 步：注册并新建项目
1. 打开 <https://supabase.com> → **Start your project**（用 GitHub 账号登录最快）
2. **New project**：
   - Name：随便，如 `prop-recorder`
   - Database Password：随便设（这个我们用不到，但必须填）
   - **Region：选 `Southeast Asia (Singapore)`**（离国内最近，速度最好）
   - 点 **Create new project**，等 1~2 分钟初始化完成

### 第 2 步：建表（复制粘贴即可）
左侧 **SQL Editor → New query**，粘贴下面整段，点 **Run**：

```sql
-- 记录表：每条记录一行，内容存 jsonb
create table if not exists public.prop_records (
  id         bigint generated always as identity primary key,
  record_id  text,
  created_at timestamptz not null default now(),
  data       jsonb not null
);

-- 打开行级安全（RLS）
alter table public.prop_records enable row level security;

-- 允许匿名读取 / 新增 / 删除（团队内部共享用；想加写入口令见文末）
create policy "anon read"   on public.prop_records for select using (true);
create policy "anon insert" on public.prop_records for insert with check (true);
create policy "anon delete" on public.prop_records for delete using (true);

-- 权限
grant usage on schema public to anon;
grant select, insert, delete on public.prop_records to anon;
```

### 第 3 步：拿到两个值
Supabase 控制台左下角 **Settings → API Keys**（新界面已经没有单独的 "API" 页了；
也可以点项目首页的 **Connect** 按钮，弹窗第一行就是 Project URL）。

- **Project URL** —— 样子是：`https://你的项目ID.supabase.co`
  例：`https://abcdefghijklmnopqrst.supabase.co`
  （`https://` + 一串小写字母数字 + `.supabase.co`；**结尾不要带斜杠**）
  - 也可以从浏览器地址栏看出来：`https://supabase.com/dashboard/project/abcdefghijklmnopqrst`
    里的 `abcdefghijklmnopqrst` 就是它
  - ⚠️ 别填成：控制台地址（`supabase.com/dashboard/...`）、数据库连接串（`postgresql://...`）、或 `db.xxx.supabase.co`
- **公开 key**（页面上叫 **publishable key** 或旧版的 **anon public**）
  - 新版：`sb_publishable_...`（短串）
  - 旧版：`eyJ...`（很长的 JWT）
  - 两种都支持，整串复制粘贴即可
  - ⚠️ **千万别用 `sb_secret_...` / `service_role`**，那是管理员密钥，会绕过所有权限限制

### 第 4 步：填进 `config.js`
在 GitHub 仓库里点开 `config.js` → 右上角**铅笔**编辑 → 填好这两行 → **Commit changes**：

```js
window.PROP_CONFIG = {
  SUPABASE_URL: 'https://abcdefghijklmn.supabase.co',
  SUPABASE_ANON_KEY: 'eyJhbGciOi...（粘贴你复制的那一串）',
  TABLE: 'prop_records'
};
```

等 Pages 自动重新部署（约 1 分钟），刷新页面：
- 右上角变成 **「云端共享 · N 条」** = 成功
- 让队友打开同一个网址，录一条 → 你点「刷新」就能看到
- 若显示 **「云端共享（连接异常）」** 并有红条，就是 URL/Key 填错了，或表/策略没建好

### 免费额度
- 数据库 500MB、2 个项目，几个人的测试记录完全够用
- 免费项目 **7 天没有任何请求会进入休眠**；在 Supabase 控制台点一下即可唤醒，数据不会丢

### ⚠️ 关于权限（重要）
`anon key` 是**公开**的（网页源码里就能看到，这是 Supabase 的设计），所以上面的策略等于
**"拿到网页链接的人都能新增/删除记录"**，类似"共享链接的在线文档"。
- 想更严格（只有知道口令的人能写）可以再加一段 RLS 策略，用 `data->>'code'` 校验写入口令，需要的话告诉我，我给你改。
- 无论用哪种模式，都建议定期点「导出 Excel / JSON 备份」留底。

---

## 免责声明

内置电机 / 电调参数为公开资料整理的参考值，可能与厂家规格有出入，请以官方数据为准。
