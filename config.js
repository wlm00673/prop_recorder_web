/* ============================================================
 * 螺旋桨记录器 · 配置
 * ------------------------------------------------------------
 * 不填（保持空字符串）= 数据保存在本机浏览器（每个人一份）
 * 填了 Supabase 两项   = 所有人共享同一份云端数据
 *
 * 配置步骤见 README.md「让所有人共享同一份数据」一节。
 * ============================================================ */
window.PROP_CONFIG = {
  /* ── 1) Project URL ─────────────────────────────────────────
   * 样子：https://你的项目ID.supabase.co
   * 例：  https://abcdefghijklmnopqrst.supabase.co
   * 在哪找：Supabase 控制台 → 左下 Settings → API Keys
   *        （新界面没有单独的 "API" 页；也可点首页的 Connect 按钮，第一行就是它）
   * 注意：结尾不要带斜杠，也不要填成控制台地址 supabase.com/dashboard/...
   * ────────────────────────────────────────────────────────── */
  SUPABASE_URL: 'https://idwyzdtwwxmaaaqwkwuo.supabase.co',

  /* ── 2) 公开 key（可安全放在网页里）────────────────────────
   * 新版：sb_publishable_...（短串）
   * 旧版：eyJ...（很长的 JWT，叫 anon public）
   * 两种都支持，直接整串粘贴。千万别填 sb_secret_... / service_role（那是管理员密钥）
   * ────────────────────────────────────────────────────────── */
  SUPABASE_ANON_KEY: '',

  /* 数据表名（按 supabase_setup.sql 建的表就叫 prop_records） */
  TABLE: 'prop_records'
};
