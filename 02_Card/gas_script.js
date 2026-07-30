/**
 * 阿福教練 - 流日能量全自動 Email 發送腳本 (GAS)
 * 請將此腳本貼入 Google 試算表的「擴充功能」->「Apps Script」中
 */

// 1. 業務員簽名檔與個人化設定 (請修改為您的資訊)
const AGENT_CONFIG = {
  senderName: "阿福教練｜流日能量關懷", // 客戶收件匣看到的寄件者顯示名稱
  name: "阿福教練",
  dept: "富邦人壽 嘉義興業通訊處",
  phone: "0988-777-666",
  lineUrl: "https://line.me/ti/p/~afu168"
};

// 2. 每日自動執行觸發點 (請手動執行此函式一次即可完成安裝)
function setupDailyTrigger() {
  // 清除舊的觸發器
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  // 設定每天早上 08:00 自動執行 dailyCheckAndSend
  ScriptApp.newTrigger('dailyCheckAndSend')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
    
  Logger.log("✅ 每日 08:00 自動發信觸發器安裝成功！");
}

// 3. 核心處理邏輯：檢查試算表並發信
function dailyCheckAndSend() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return; // 無資料

  const today = new Date();
  const tY = today.getFullYear();
  const tM = today.getMonth() + 1;
  const tD = today.getDate();
  const todayStr = `${tY}-${tM}-${tD}`;

  // 走訪試算表每一行資料 (第 2 行開始)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[0];        // A: 客戶姓名
    const nickname = row[1];    // B: 客戶暱稱
    const company = row[2];     // C: 公司
    const dept = row[3];        // D: 服務通訊處
    const phone = row[4];       // E: 手機
    const email = row[5];       // F: Email
    const region = row[6];      // G: 區域 (嘉義,台南,高雄,屏東...)
    const bM = parseInt(row[7]);// H: 生日月
    const bD = parseInt(row[8]);// I: 生日日
    const note = row[9];        // J: 備註
    let status = row[10];       // K: 自動發送狀態
    let progress = parseInt(row[11]) || 0; // L: 當前發送進度
    let lastDate = row[12];     // M: 最後發送日期

    if (!name || !email || !bM || !bD) continue;
    if (status === "暫停" || status === "已完成") continue;
    if (lastDate === todayStr) continue; // 今天已發送過

    const displayName = (nickname && String(nickname).trim()) ? String(nickname).trim() : name;

    // 檢查狀況 1：今天是否生日？ -> 寄發第 0 則生日卡
    if (tM === bM && tD === bD && progress === 0) {
      const py = getPY(tY, tM, bM, bD);
      const subject = `🎂 祝 ${displayName} 生日快樂！給您的專屬流年祝福與能量禮物 🌟`;
      const html = getBirthdayHtml(displayName, py, AGENT_CONFIG);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: html,
        name: AGENT_CONFIG.senderName || AGENT_CONFIG.name
      });
      
      // 更新試算表狀態
      sheet.getRange(i + 1, 11).setValue("發送中");
      sheet.getRange(i + 1, 12).setValue(1); // 下一次進度為 Day 1
      sheet.getRange(i + 1, 13).setValue(todayStr);
      continue;
    }

    // 檢查狀況 2：連續 7 天流日發送 (進度 1~7)
    if (progress >= 1 && progress <= 7) {
      const py = getPY(tY, tM, bM, bD);
      const pm = getPM(py, tM);
      const pd = getPD(pm, tD);
      const flow = workflowData[pd];

      const subject = `🌟 【流日${pd}・${flow.short}】給 ${displayName} 的第 ${progress} 天能量行動提案：${flow.subtitle}`;
      const html = getFlowDayHtml(progress, displayName, pd, flow, AGENT_CONFIG);

      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: html,
        name: AGENT_CONFIG.senderName || AGENT_CONFIG.name
      });

      // 更新進度
      const nextProgress = progress + 1;
      const nextStatus = nextProgress > 7 ? "已完成" : "發送中";
      
      sheet.getRange(i + 1, 11).setValue(nextStatus);
      sheet.getRange(i + 1, 12).setValue(nextProgress);
      sheet.getRange(i + 1, 13).setValue(todayStr);
    }
  }
}

// 4. 生命靈數計算
function dr(n) { return n === 0 ? 0 : (n - 1) % 9 + 1; }
function getPY(tY, tM, bM, bD) { return dr((tM < bM ? tY - 1 : tY) + bM + bD); }
function getPM(pY, cM) { return dr(pY + cM); }
function getPD(pM, cD) { return dr(pM + cD); }

// 5. 流日資料
const workflowData = {
  1: { name:"破局與開創日", subtitle:"舒適圈粉碎", short:"破局", tagline:"不用等待許可，今天你就是自己生命中最強大的貴人！", quote:"「在猶豫中等待機會，不如用勇敢劈開第一道曙光！今天不設限，我的獨立與果斷，就是最強大的開創磁場。」", tip:"眼前的機會絕不猶豫，今天必須拿出勇氣、獨立出擊，播下最勇敢的種子。", checklist:["今天有沒有做一件平時最不敢做的勇敢嘗試？","面對機會，有沒有當機立斷「立刻行動」？"] },
  2: { name:"借力與傾聽日", subtitle:"溫柔凝聚力", short:"借力", tagline:"收起鋒芒、張開雙耳。學會借力，才是最高明的大格局。", quote:"「平時靠實力，今天靠魅力與聽力。縮小自己、放大對方，用最高級的溫柔凝聚全宇宙的人脈複利！」", tip:"今天最強大的力量，是靜靜傾聽對方說的每一句話。", checklist:["今天有沒有忍住說話衝動，好好把對方的話聽完？","有沒有真誠表達感謝，讓對方感受到被重視？"] },
  3: { name:"表達與促成日", subtitle:"創意渲染力", short:"促成", tagline:"用創意的糖衣，包裹理性的專業；你的快樂就是最強大的促成力！", quote:"「不嚴肅、不糾結！今天我是自己舞台的發表者，用最陽光的笑容去利他，把熱情化為現實的黃金成果。」", tip:"今天的你，就是最有感染力的那個人。展現快樂，就能感染身邊所有人！", checklist:["今天有沒有釋放出樂觀、喜悅的正向能量？","面對卡關，有沒有用幽默或創意的方式嘗試突破？"] },
  4: { name:"務實與扎根日", subtitle:"信任落地方案", short:"扎根", tagline:"不給模糊承諾，答應的事精準做到；用無懈可擊的誠信築起護城河！", quote:"「今天不浮誇、不浮躁！用最扎實的行動與誠信，將善緣化為最讓人安心的長久關係。」", tip:"今天做的每一件踏實小事，都是未來最穩固的基石。", checklist:["今天答應別人的事，有沒有精準做到，不留模糊空間？","待辦事項和重要環境，今天是否有確實整理？"] },
  5: { name:"跨界與突破日", subtitle:"打破舒適圈", short:"擴張", tagline:"守成就是最大的風險！帶著破局之翼，讓遠大願景飛越一切限制。", quote:"「拒絕無效瞎忙，大膽打破現狀！今天我不設限，我的勇敢與跨界，就是最強大的擴張磁場。」", tip:"今天做一件讓自己稍微不舒服但充滿可能性的新嘗試！", checklist:["今天有沒有主動走向新環境或接觸新朋友？","有沒有拒絕一件低效的事，把時間留給真正重要的目標？"] },
  6: { name:"擔當與修復日", subtitle:"智慧與大愛", short:"修復", tagline:"用一肩扛起責任的肩膀修復卡關；你的擔當，就是最療癒的磁場！", quote:"「不怕麻煩、主動站出來！用超越期待的責任心，將眼前的困境轉化為彼此一輩子的深厚信任。」", tip:"今天用愛與責任修復所有卡關，你的擔當就是最美的力量。", checklist:["面對難題，有沒有第一時間勇敢說「我來處理」？","今天有沒有讓身邊焦慮的人，因自己而重獲平靜？"] },
  7: { name:"底層邏輯與請益日", subtitle:"軍師智慧", short:"貴人", tagline:"用清晰的眼光看透真相；以謙遜請益，借力生命中的頂級貴人。", quote:"「今天不盲從、不浮躁！退回大後方深度整理，我的清明與深度，就是吸引最強高階貴人的磁場。」", tip:"今天把時間留給深度思考，一個清晰的決策，勝過百個倉促的行動。", checklist:["今天有沒有用邏輯與理性做重要決定？","有沒有主動向比自己厲害的人學習或請益？"] },
  8: { name:"主導與顯化日", subtitle:"黃金收割", short:"收割", tagline:"理直氣壯展現自己！換上戰袍，完成最高規格的豐盛顯化。", quote:"「收起不好意思的心態！今天用最強大的領袖氣場與利他心，理直氣壯地將善緣轉化為實質的豐盛果實。」", tip:"今天換上自信的戰袍，大聲說出自己的價值！你值得擁有最好的。", checklist:["今天有沒有展現出自信的氣場，拿回自己的主導權？","有沒有果斷推動一件重要的事，讓它真正落地？"] },
  9: { name:"圓滿與斷捨離日", subtitle:"奇蹟藍圖", short:"圓滿", tagline:"帶著豁達感恩完美收網；果斷清空記憶體，預約下一個奇蹟！", quote:"「向舊世界優雅告別！今天我不糾結、不留戀，用豁達與純粹利他，迎接下一個大循環的奇蹟新局。」", tip:"今天帶著感恩完美結尾，然後勇敢放手，讓更好的事物到來。", checklist:["今天有沒有把卡關或懸而未決的事，追蹤到底、完美結案？","有沒有勇敢放下一個消耗自己的事物，清空大腦記憶體？"] }
};

const yearKw = ["開創播種","耐心扎根","展現交流","穩定建設","變動拓展","責任服務","學習沉澱","豐收成就","總結重整"];
const yearDescs = {
  1:{ d:"嶄新的能量週期正在為您開啟！這是勇於行動、播下夢想種子的一年，每一個大膽的第一步，都將化為未來豐收的基礎。", w:"願您今年以無懼之心迎向每一個新機遇，創造屬於您的精彩人生！" },
  2:{ d:"這是深耕關係、積累信任的一年。溫柔而堅定的力量，正在為生命最重要的一切打下最穩固的根基。", w:"願您今年在每一段珍貴關係中感受到溫暖與力量，靜待美好的綻放！" },
  3:{ d:"您的才華與魅力正等待被全世界看見！這是表達自我、廣結善緣的一年，天生的光芒將感染身邊每一個人。", w:"願您今年盡情展現最閃耀的自己，用歡笑與熱情迎接人生每一個高光時刻！" },
  4:{ d:"腳踏實地的力量，正是今年最大的禮物。這是築夢踏實、打造長久成就的一年，每一份付出都將留下最深實的印記。", w:"願您今年將每一個計畫化為堅實的成果，用勤勉與誠信創造令人欽佩的人生成就！" },
  5:{ d:"自由與突破是今年最強大的能量！這是勇敢跨越舒適圈、擁抱無限可能的一年，改變正是最好的禮物。", w:"願您今年乘著改變的翅膀，勇敢飛向更廣闊的天地，讓生命充滿驚喜與精彩！" },
  6:{ d:"今年，深情與擔當將成為身邊人最溫暖的依靠。這是深耕關係、收穫深厚情誼的一年，愛是最強大的能量。", w:"願您今年在給予與被愛中感受到最豐盛的幸福，您的善良是世界最美的禮物！" },
  7:{ d:"智慧的大門正為您敞開！這是洞察真理、提升生命格局的一年，每一次靜思都是靈魂最深的升華。", w:"願您今年在知識與智慧中找到最深的寧靜與力量，成就更深刻而豐富的自己！" },
  8:{ d:"多年的努力正在迎來最美好的收穫！這是豐盛顯化、實現夢想的一年，豐收與成就正向您奔來。", w:"願您今年大豐收，所有的付出都得到最圓滿的回報，盡情享受您應得的一切美好！" },
  9:{ d:"圓滿與智慧是今年最動人的能量。這是優雅告別舊循環、以感恩之心迎向嶄新篇章的一年。", w:"願您今年帶著滿滿的感恩與智慧，圓滿完成每一個美好的故事，優雅迎向生命的嶄新精彩！" }
};

// 6. HTML 信件範本生成
function getBirthdayHtml(cn, py, agent) {
  const theme = yearKw[py-1];
  const yd = yearDescs[py];
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:20px;background-color:#f1f5f9;font-family:'Microsoft JhengHei',sans-serif;"><div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 10px 25px rgba(0,0,0,0.08);"><div style="background:linear-gradient(135deg, #4f46e5 0%, #7c3aed 50%, #c026d3 100%);padding:40px 30px;text-align:center;color:#ffffff;"><div style="font-size:48px;margin-bottom:10px;">🎂</div><h1 style="margin:0;font-size:26px;font-weight:800;">祝 ${cn} 生日快樂！</h1><p style="margin:8px 0 0 0;font-size:14px;opacity:0.9;">獻上最真摯的祝福 🌟</p></div><div style="padding:35px 30px;color:#334155;line-height:1.8;"><p style="font-size:16px;font-weight:bold;color:#1e293b;margin-top:0;">親愛的 ${cn} 您好：</p><p style="font-size:15px;">祝您生日快樂、身體健康、萬事順心！</p><div style="background:#f8fafc;border-left:4px solid #7c3aed;padding:20px;border-radius:8px;margin:25px 0;"><div style="font-size:12px;font-weight:bold;color:#7c3aed;text-transform:uppercase;margin-bottom:6px;">✨ 您的年度流年能量洞察</div><div style="font-size:18px;font-weight:bold;color:#1e293b;margin-bottom:10px;">今年您進入「${theme}之年」（流年 ${py}）</div><p style="font-size:14px;color:#475569;margin:0 0 10px 0;">${yd.d}</p><p style="font-size:14px;font-weight:bold;color:#6d28d9;margin:0;">💫 ${yd.w}</p></div><p style="font-size:15px;">感謝您一直以來的信任與支持！</p></div><div style="background:#0f172a;padding:25px 30px;color:#f8fafc;border-top:3px solid #fbbf24;"><table style="width:100%;border-collapse:collapse;"><tr><td><div style="font-size:18px;font-weight:bold;color:#fbbf24;margin-bottom:4px;">${agent.name}</div><div style="font-size:13px;color:#94a3b8;margin-bottom:8px;">${agent.dept}</div><div style="font-size:13px;color:#cbd5e1;">📱 手機：${agent.phone}</div></td><td style="text-align:right;vertical-align:bottom;"><a href="${agent.lineUrl}" target="_blank" style="display:inline-block;background:#22c55e;color:#ffffff;text-decoration:none;padding:10px 18px;border-radius:30px;font-size:13px;font-weight:bold;">LINE 諮詢 / 預約</a></td></tr></table></div></div></body></html>`;
}

function getFlowDayHtml(n, cn, pd, flow, agent) {
  const cl = (flow.checklist||[]).map(c=>`<li style="margin-bottom:6px;">${c}</li>`).join('');
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:20px;background-color:#f1f5f9;font-family:'Microsoft JhengHei',sans-serif;"><div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:20px;overflow:hidden;box-shadow:0 10px 25px rgba(0,0,0,0.08);"><div style="background:#1e293b;padding:14px 25px;color:#cbd5e1;font-size:13px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #334155;"><span>📅 第 ${n} 則・流日能量關懷訊息</span><span style="background:#f59e0b;color:#000000;padding:2px 10px;border-radius:12px;font-weight:bold;font-size:12px;">流日 ${pd}・${flow.short}</span></div><div style="background:linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%);padding:35px 30px;color:#ffffff;"><div style="font-size:13px;color:#fbbf24;font-weight:bold;margin-bottom:6px;">【${flow.name}】${flow.subtitle}</div><h1 style="margin:0 0 12px 0;font-size:22px;font-weight:800;color:#ffffff;line-height:1.4;">${flow.tagline}</h1><div style="font-size:14px;color:#cbd5e1;font-style:italic;line-height:1.6;border-left:3px solid #fbbf24;padding-left:12px;margin-top:10px;">${flow.quote}</div></div><div style="padding:30px;color:#334155;line-height:1.8;"><p style="font-size:15px;margin-top:0;">嗨 <strong>${cn}</strong> 您好：</p><div style="background:#fffbeb;border:1px solid #fef3c7;padding:18px;border-radius:12px;margin:20px 0;"><div style="font-size:13px;font-weight:bold;color:#b45309;margin-bottom:4px;">💡 今日能量提醒</div><div style="font-size:14px;color:#78350f;">${flow.tip}</div></div><div style="background:#f0fdf4;border:1px solid #dcfce7;padding:18px;border-radius:12px;margin-bottom:25px;"><div style="font-size:13px;font-weight:bold;color:#15803d;margin-bottom:8px;">🧭 夜間 21:00 執行檢視提案</div><ul style="margin:0;padding-left:20px;font-size:14px;color:#166534;">${cl}</ul></div><p style="font-size:14px;color:#64748b;margin-bottom:0;">祝福您今天擁有順心充實的一天！ 😊</p></div><div style="background:#0f172a;padding:25px 30px;color:#f8fafc;border-top:3px solid #fbbf24;"><table style="width:100%;border-collapse:collapse;"><tr><td><div style="font-size:18px;font-weight:bold;color:#fbbf24;margin-bottom:4px;">${agent.name}</div><div style="font-size:13px;color:#94a3b8;margin-bottom:8px;">${agent.dept}</div><div style="font-size:13px;color:#cbd5e1;">📱 手機：${agent.phone}</div></td><td style="text-align:right;vertical-align:bottom;"><a href="${agent.lineUrl}" target="_blank" style="display:inline-block;background:#22c55e;color:#ffffff;text-decoration:none;padding:10px 18px;border-radius:30px;font-size:13px;font-weight:bold;">LINE 諮詢 / 預約</a></td></tr></table></div></div></body></html>`;
}
