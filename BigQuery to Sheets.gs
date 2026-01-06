/**
 * BigQueryからデータを取得してスプレッドシートに書き込むスクリプト
 * 
 * 【データ取得】
 * dataKoushin() を実行
 */



/**
 * GCPプロジェクトID設定
 * 初回のみ実行してプロジェクトIDを設定
 */
function projectIdSettei() {
  const projectId = 'pltfrm-prod'; // ここにGCPプロジェクトIDを記載
  
  try {
    PropertiesService.getScriptProperties().setProperty('BQ_PROJECT_ID', projectId);
    Logger.log('✓ GCPプロジェクトIDを設定しました: ' + projectId);
    Logger.log('これで dataKoushin() を実行できます');
  } catch (error) {
    Logger.log('✗ エラー: プロジェクトIDの設定に失敗しました');
    Logger.log(error.toString());
  }
}

/**
 * データ更新
 * BigQueryから店舗リストとユーザーリストを取得してシートに書き込む
 */
function dataKoushin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log('=== データ取得を開始します ===');
    
    // プロジェクトIDの確認
    const projectId = PropertiesService.getScriptProperties().getProperty('BQ_PROJECT_ID');
    if (!projectId) {
      Logger.log('✗ エラー: GCPプロジェクトIDが設定されていません');
      Logger.log('projectIdSettei() を実行して設定してください');
      return;
    }
    
    Logger.log('使用するプロジェクトID: ' + projectId);
    
    // 店舗リスト取得
    Logger.log('店舗リストを取得中...');
    const tenpoData = tenpoListSakusei(ss);
    Logger.log('✓ 店舗リストの取得が完了しました');
    
    // ユーザーリスト取得
    Logger.log('ユーザーリストを取得中...');
    const userData = userListSakusei(ss);
    Logger.log('✓ ユーザーリストの取得が完了しました');
    
    // リージョン別シート作成
    Logger.log('リージョン別シートを作成中...');
    regionSheetSakusei(ss, tenpoData, userData);
    Logger.log('✓ リージョン別シートの作成が完了しました');
    
    Logger.log('=== すべてのデータ取得が完了しました! ===');
    
  } catch (error) {
    Logger.log('✗ エラーが発生しました: ' + error.toString());
    throw error;
  }
}

/**
 * 店舗リストシート作成（ソース列付きマージ対応）
 * - auto: BigQueryから取得したデータ
 * - manual: 手動追加したデータ（E列に'manual'と入力）
 * - BigQuery更新時はautoのみ上書き、manualは維持
 * - BigQueryに同じ店舗コードが登録されたらmanualを自動削除
 * ※列構造: A:店舗コード, B:店舗名, C:リージョン, D:店舗タイプ, E:ソース（参照先は変更なし）
 */
function tenpoListSakusei(ss) {
  const sheetName = '店舗リスト';
  
  const query = `
    SELECT
      s1.code,
      name as shopName,
      s1.region,
      shop_type as shopType
    FROM
      \`exment.SHOPS\` s1
    WHERE brand = 'crisp'
      AND s1.code LIKE 'CSW%'
    ORDER BY code
  `;
  
  const data = bigQueryJikkou(query);
  
  let sheet = ss.getSheetByName(sheetName);
  let manualRows = [];
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('新規シート「' + sheetName + '」を作成しました');
  } else {
    // 既存のmanualデータを保持
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existingData = sheet.getRange(1, 1, lastRow, 5).getValues();
      const headers = existingData[0];
      const sourceIndex = headers.indexOf('ソース');
      
      if (sourceIndex !== -1) {
        // BigQueryの店舗コード一覧を作成
        const autoCodeSet = new Set((data || []).map(row => row.code));
        
        for (let i = 1; i < existingData.length; i++) {
          const row = existingData[i];
          const source = row[sourceIndex];
          const code = row[0]; // A列=店舗コード
          
          if (source === 'manual') {
            if (autoCodeSet.has(code)) {
              Logger.log(`✓ 手動データ「${code}」はBigQueryに登録済みのため削除`);
            } else {
              manualRows.push(row);
            }
          }
        }
        Logger.log(`手動データ: ${manualRows.length}件を維持`);
      }
    }
    sheet.clear();
    Logger.log('既存シート「' + sheetName + '」をクリアしました');
  }
  
  // ヘッダー（E列にソース追加）
  const headers = [['店舗コード', '店舗名', 'リージョン', '店舗タイプ', 'ソース']];
  sheet.getRange(1, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // BigQueryデータ（auto）
  const autoRows = (data || []).map(row => [
    row.code || '',
    row.shopName || '',
    row.region || '',
    row.shopType || '',
    'auto'
  ]);
  
  // autoとmanualを統合
  const allRows = [...autoRows, ...manualRows];
  
  if (allRows.length > 0) {
    sheet.getRange(2, 1, allRows.length, 5).setValues(allRows);
    Logger.log(`店舗データ: auto=${autoRows.length}件, manual=${manualRows.length}件`);
  }
  
  // 列幅設定
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 80);
  sheet.setFrozenRows(1);
  
  // 返り値：元の形式を維持（manualも含めたオブジェクト配列）
  // regionSheetSakuseiでの参照用にregionプロパティを含める
  const allData = (data || []).map(row => ({
    code: row.code,
    shopName: row.shopName,
    region: row.region,
    shopType: row.shopType
  }));
  // manualデータも追加
  manualRows.forEach(row => {
    allData.push({
      code: row[0],
      shopName: row[1],
      region: row[2],
      shopType: row[3]
    });
  });
  
  return allData;
}

/**
 * ユーザーリストシート作成（ソース列付きマージ対応）
 * - auto: BigQueryから取得したデータ
 * - manual: 手動追加したデータ（最終列に'manual'と入力）
 * - BigQuery更新時はautoのみ上書き、manualは維持
 * - BigQueryに同じIDが登録されたらmanualを自動削除
 * ※列構造は変更なし（既存カラム + 最終列にsource追加）
 */
function userListSakusei(ss) {
  const sheetName = 'ユーザーリスト';
  
  const query = `
    SELECT * 
    FROM \`jinjer.EMPLOYEES\`
    WHERE enrollmentClassification = '在籍'
      AND REGEXP_CONTAINS(departmentName, r'^Region [A-Z]$')
      AND ID != '99998'
  `;
  
  const data = bigQueryJikkou(query);
  
  let sheet = ss.getSheetByName(sheetName);
  let manualRows = [];
  let existingHeaders = null;
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('新規シート「' + sheetName + '」を作成しました');
  } else {
    // 既存のmanualデータを保持
    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) {
      existingHeaders = existingData[0];
      const sourceIndex = existingHeaders.indexOf('source');
      const idIndex = existingHeaders.indexOf('id');
      
      if (sourceIndex !== -1 && idIndex !== -1) {
        // BigQueryのID一覧を作成
        const autoIdSet = new Set((data || []).map(row => String(row.id)));
        
        for (let i = 1; i < existingData.length; i++) {
          const row = existingData[i];
          const source = row[sourceIndex];
          const id = String(row[idIndex]);
          
          if (source === 'manual') {
            if (autoIdSet.has(id)) {
              Logger.log(`✓ 手動データ「${id}」はBigQueryに登録済みのため削除`);
            } else {
              manualRows.push(row);
            }
          }
        }
        Logger.log(`手動データ: ${manualRows.length}件を維持`);
      }
    }
    sheet.clear();
    Logger.log('既存シート「' + sheetName + '」をクリアしました');
  }
  
  if (data && data.length > 0) {
    // ヘッダー（最終列にsource追加）
    const baseHeaders = Object.keys(data[0]);
    const headers = [...baseHeaders, 'source'];
    
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    
    // BigQueryデータ（auto）
    const autoRows = data.map(row => [
      ...baseHeaders.map(header => row[header] || ''),
      'auto'
    ]);
    
    // manualデータを新しいヘッダーに合わせる
    const adjustedManualRows = manualRows.map(oldRow => {
      const newRow = headers.map(h => '');
      if (existingHeaders) {
        headers.forEach((h, newIdx) => {
          const oldIdx = existingHeaders.indexOf(h);
          if (oldIdx !== -1) {
            newRow[newIdx] = oldRow[oldIdx];
          }
        });
      }
      return newRow;
    });
    
    // autoとmanualを統合
    const allRows = [...autoRows, ...adjustedManualRows];
    
    if (allRows.length > 0) {
      sheet.getRange(2, 1, allRows.length, headers.length).setValues(allRows);
      Logger.log(`ユーザーデータ: auto=${autoRows.length}件, manual=${adjustedManualRows.length}件`);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    sheet.setFrozenRows(1);
  }
  
  // 返り値：元の形式を維持（BigQueryデータそのまま）
  return data;
}

/**
 * リージョン別シート作成
 * 各ユーザーごとにシートを作成し、担当者と対象店舗を表示
 * シート名：「フルネーム リージョンアルファベット」形式
 * 
 * 統合ロジック：
 * - フルネーム：ユーザーリストのC列 + B列を結合（名 + 姓）
 * - 担当者列：フルネームのみ（IDなし）
 * - 対象店舗：店舗リストのA列 + B列
 * - マッチング：店舗リストのC列とユーザーリストのU列で合致
 */
function regionSheetSakusei(ss, tenpoData, userData) {
  if (!tenpoData || !userData || tenpoData.length === 0 || userData.length === 0) {
    Logger.log('店舗データまたはユーザーデータが存在しません');
    return;
  }
  
  // ユーザーリストシートから実際のデータを読み取る
  const userSheet = ss.getSheetByName('ユーザーリスト');
  if (!userSheet) {
    Logger.log('✗ エラー: ユーザーリストシートが見つかりません');
    return;
  }
  
  const userSheetData = userSheet.getDataRange().getValues();
  const userHeaders = userSheetData[0];
  const userRows = userSheetData.slice(1);
  
  Logger.log('=== ユーザーリストシート構造の確認 ===');
  Logger.log('全カラム数: ' + userHeaders.length);
  Logger.log('A列(0): ' + userHeaders[0]);
  Logger.log('B列(1): ' + userHeaders[1]);
  Logger.log('C列(2): ' + userHeaders[2]);
  Logger.log('U列(20): ' + userHeaders[20]);
  
  // カラムインデックスを取得
  const idIndex = 0; // A列 (0始まりなので0) - ID
  const lastNameIndex = 1; // B列 (0始まりなので1) - 姓
  const firstNameIndex = 2; // C列 (0始まりなので2) - 名
  const deptIndex = 20; // U列 (0始まりなので20) - リージョン
  
  // ユーザーデータをオブジェクトに変換
  const userList = userRows.map((row, rowIndex) => {
    const lastName = row[lastNameIndex] || '';
    const firstName = row[firstNameIndex] || '';
    const fullName = `${firstName}${lastName}`; // C列 + B列でフルネーム（名 + 姓）
    
    const user = {
      id: row[idIndex] || '',
      lastName: lastName,
      firstName: firstName,
      fullName: fullName,
      department: row[deptIndex] || ''
    };
    
    // 最初の3ユーザーをログ出力
    if (rowIndex < 3) {
      Logger.log(`ユーザー${rowIndex + 1}: ID=${user.id}, FirstName=${user.firstName}, LastName=${user.lastName}, FullName=${user.fullName}, Dept=${user.department}`);
    }
    
    return user;
  });
  
  // リージョンのリストを取得（店舗データとユーザーデータの両方から、重複なし）
  // 正規化: すべて "Region X" 形式に統一
  const regionSet = new Set();
  
  // 店舗データからリージョンを抽出（"A" → "Region A" に正規化）
  tenpoData.forEach(tenpo => {
    if (tenpo.region) {
      const r = tenpo.region.toString().trim();
      // 単一アルファベットの場合は "Region X" 形式に変換
      if (/^[A-Z]$/i.test(r)) {
        regionSet.add(`Region ${r.toUpperCase()}`);
      } else {
        regionSet.add(r);
      }
    }
  });
  
  // ユーザーデータからもリージョンを抽出（店舗がなくてもユーザーがいれば対応）
  userList.forEach(user => {
    const match = (user.department || '').match(/^Region\s*([A-Z])$/i);
    if (match) {
      regionSet.add(`Region ${match[1].toUpperCase()}`);
    }
  });
  
  const regions = Array.from(regionSet).sort();
  Logger.log('検出されたリージョン（店舗+ユーザー統合）: ' + regions.join(', '));
  
  // ユーザーごとにシートを作成
  userList.forEach(user => {
    const userFullName = user.fullName || user.id || '名前なし';
    const userDept = user.department || '';
    
    // ユーザーの部署名からリージョンを直接抽出
    const deptMatch = userDept.match(/^Region\s*([A-Z])$/i);
    if (!deptMatch) {
      Logger.log(`⚠ ユーザー ${userFullName} はリージョン形式に一致しません（部署: ${userDept}）`);
      return;
    }
    
    const regionLetter = deptMatch[1].toUpperCase();
    const userRegion = `Region ${regionLetter}`;
    
    // シート名：「フルネーム リージョンアルファベット」
    const sheetName = `${userFullName} ${regionLetter}`;
    
    Logger.log(`シート作成: ${sheetName} (フルネーム: ${userFullName}, リージョン: ${userRegion})`);
    
    // このユーザーが担当する店舗を抽出（リージョン形式の違いに対応: "A" または "Region A"）
    const userTenpo = tenpoData.filter(tenpo => {
      const tenpoRegion = (tenpo.region || '').toString().trim();
      return tenpoRegion === regionLetter || 
             tenpoRegion === userRegion ||
             tenpoRegion.toUpperCase() === regionLetter;
    });
    
    Logger.log(`  ${sheetName}: 店舗数=${userTenpo.length}`);
    
    // シートの取得または作成
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log('  新規シート「' + sheetName + '」を作成しました');
    } else {
      sheet.clear();
      Logger.log('  既存シート「' + sheetName + '」をクリアしました');
    }
    
    // ヘッダー作成
    const headers = [['担当者', '対象店舗']];
    sheet.getRange(1, 1, 1, 2)
      .setValues(headers)
      .setFontWeight('bold')
      .setBackground('#f4b400')
      .setFontColor('#ffffff');
    
    // データ行作成
    const rows = [];
    
    if (userTenpo.length > 0) {
      userTenpo.forEach(tenpo => {
        // 担当者：フルネームのみ（C列+B列 = 名+姓）
        const tantosha = user.fullName || user.id || '';
        
        // 対象店舗：A列（code）スペース B列（shopName）
        const tenpoCode = tenpo.code || '';
        const tenpoName = tenpo.shopName || '';
        const tenpoDisplay = `${tenpoCode} ${tenpoName}`;
        
        rows.push([tantosha, tenpoDisplay]);
      });
    } else {
      // 店舗がない場合はプレースホルダーを追加
      rows.push([user.fullName || user.id || '', '（店舗データなし - 登録後に自動更新されます）']);
    }
    
    // データ書き込み
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    if (userTenpo.length > 0) {
      Logger.log(`  ${sheetName}: ${rows.length}件のデータを書き込みました`);
    } else {
      Logger.log(`  ${sheetName}: 店舗データなし（プレースホルダーを追加）`);
    }
    
    // 列幅設定
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 350);
    sheet.setFrozenRows(1);
  });
}

/**
 * BigQueryクエリ実行
 */
function bigQueryJikkou(query) {
  try {
    Logger.log('BigQueryクエリを実行中...');
    
    const projectId = PropertiesService.getScriptProperties().getProperty('BQ_PROJECT_ID');
    if (!projectId) {
      throw new Error('GCPプロジェクトIDが設定されていません');
    }
    
    const request = {
      query: query,
      useLegacySql: false
    };
    
    const queryResults = BigQuery.Jobs.query(request, projectId);
    const jobId = queryResults.jobReference.jobId;
    
    let rows = queryResults.rows;
    let pageToken = queryResults.pageToken;
    
    while (pageToken) {
      const results = BigQuery.Jobs.getQueryResults(projectId, jobId, {
        pageToken: pageToken
      });
      rows = rows.concat(results.rows);
      pageToken = results.pageToken;
    }
    
    const headers = queryResults.schema.fields.map(field => field.name);
    const data = rows.map(row => {
      const obj = {};
      row.f.forEach((cell, index) => {
        obj[headers[index]] = cell.v;
      });
      return obj;
    });
    
    Logger.log('取得件数: ' + data.length + '件');
    return data;
    
  } catch (error) {
    Logger.log('✗ BigQueryエラー: ' + error.toString());
    throw error;
  }
}
