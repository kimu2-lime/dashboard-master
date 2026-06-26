// ============================================================
//  デバッグ用：直営店目標が読めているか診断
//  GASエディタで testDirectTargets を実行 → ログを確認
// ============================================================
function testDirectTargets() {
  const ss = getSS('SS_DIRECT');
  Logger.log('SS_DIRECT タブ一覧: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(' / '));

  const t = getDirectTargets(null);
  const stores = Object.keys(t);
  Logger.log('────────────');
  Logger.log('直営目標 読み取り店舗数: ' + stores.length);

  if (!stores.length) {
    Logger.log('⚠ 0店舗しか読めていません。よくある原因：');
    Logger.log('  ① タブ名が「202606」のような YYYYMM 形式になっていない（「2026/06」「6月目標」等だと読まれない）');
    Logger.log('  ② ヘッダー行（既定の所定行）に店舗名が入っていない / 列がずれている');
    Logger.log('  → 上の「タブ一覧」を見て、当月の目標タブが YYYYMM 形式か確認してください');
    return;
  }

  stores.forEach(function(s) {
    Object.keys(t[s]).forEach(function(m) {
      const x = t[s][m];
      Logger.log(s + ' [' + m + '] 目標売上=' + x.total_sales + ' / 新規=' + x.new_count
        + ' / 契約=' + x.new_contract_count + ' / 契約単価=' + x.new_contract_unit_price);
    });
  });
  Logger.log('────────────');
  Logger.log('※ 目標売上=0 の店が多い → 入力行がズレている可能性（ピーリング/オンダの新規数・契約率・単価などが所定の行に入っているか）');
  Logger.log('※ 店舗名がダッシュボードの略称と違う → ヘッダーの店舗名を略称に合わせる必要あり');
}
