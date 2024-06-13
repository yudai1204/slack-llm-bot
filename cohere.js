const model = "command-r-plus";

const BOT_MEMBER_ID =
  PropertiesService.getScriptProperties().getProperty("BOT_MEMBER_ID");
const BOT_CHANNEL_ID =
  PropertiesService.getScriptProperties().getProperty("BOT_CHANNEL_ID");
const BOT_AUTH_TOKEN =
  PropertiesService.getScriptProperties().getProperty("BOT_AUTH_TOKEN");
const CO_API_KEY =
  PropertiesService.getScriptProperties().getProperty("CO_API_KEY");

const fetchMsgsInThread = (channelId, threadTimestamp) => {
  const url = `https://slack.com/api/conversations.replies?channel=${channelId}&ts=${threadTimestamp}`;

  const headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + BOT_AUTH_TOKEN,
  };

  const options = {
    method: "GET",
    headers,
  };

  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());

  if (data.ok) {
    return data.messages;
  } else {
    throw new Error(`Failed to fetch messages in thread: ${data.error}`);
  }
};

/**
 * botに問いかけられたメッセージを取得する
 * @param {object} triggerMsg - Slackからのトリガーとなるメッセージ
 * @param {string} triggerMsg.channel - トリガーとなったメッセージが送信されたチャンネルID
 * @param {string} triggerMsg.text - トリガーとなったメッセージのテキスト
 * @param {string} [triggerMsg.thread_ts] - トリガーとなったメッセージが属するスレッドのタイムスタンプ（省略可）
 * @returns {object[]} - botに問いかけられたメッセージの配列。該当するものが無い場合は空配列を返す
 */
const fetchSlackMsgsAskedToBot = (triggerMsg) => {
  const isInThread = triggerMsg.thread_ts;
  const isMenthionedBot = triggerMsg.text.includes(BOT_MEMBER_ID);

  if (!isInThread) {
    // スレッド外の場合
    if (triggerMsg.channel === BOT_CHANNEL_ID || isMenthionedBot) {
      // botとのDMの場合か、botへのメンションがある場合は応答
      return [triggerMsg];
    } else {
      // それ以外の場合は無視
      return [];
    }
  } else {
    // スレッド内の場合
    const isMentionedNonBot =
      !isMenthionedBot && triggerMsg.text.includes("<@");
    if (isMentionedNonBot) {
      // bot以外へのメンションがある場合は無視
      return [];
    } else {
      // botへの問いかけと思われるスレッドの場合、スレッド内のすべてのメッセージを取得する
      const msgsInThread = fetchMsgsInThread(
        triggerMsg.channel,
        triggerMsg.thread_ts
      );

      const isBotInvolvedThread =
        msgsInThread.find((msg) => msg.user === BOT_MEMBER_ID) == null;
      if (isBotInvolvedThread && !isMenthionedBot) {
        // botと無関係のスレッドの場合は無視
        return [];
      } else {
        // botへの問いかけと思われるスレッド内のメッセージをすべて返す
        return msgsInThread;
      }
    }
  }
};

/**
 * メンションされたテキストから、メンション部分を除去して返す
 * @param {string} source - メンションされたテキスト
 * @returns {string} メンション部分が除去されたテキスト
 */
const trimMentionText = (source) => {
  const regex = /<.+> /g;
  return source.replace(regex, "").trim();
};

/**
 * SlackのメッセージをCohereのクエリメッセージ形式に変換する
 * @param {object[]} slackMsgs - Slackのメッセージオブジェクト群
 * @param {string} slackMsgs[].user - ユーザーID
 * @param {string} slackMsgs[].text - メッセージのテキスト
 * @returns {object[]} Cohereのクエリメッセージオブジェクト群
 */
const parseSlackMsgsToCohereQuesryMsgs = (slackMsgs) => {
  if (slackMsgs.length < 2) {
    return { message: trimMentionText(slackMsgs[slackMsgs.length - 1].text) };
  }
  // 配列の各要素を変換
  const chat_history = slackMsgs.slice(0, -1).forEach((msg) => {
    // BOT_MEMBER_IDと比較して、送信者がユーザーかアシスタントかを判断
    const role = msg.user == BOT_MEMBER_ID ? "CHATBOT" : "USER";
    // メンション部分を除去したテキストを取得
    const content = trimMentionText(msg.text);
    // 送信者の役割とテキストを含むメッセージオブジェクトを返す
    return {
      role: role,
      message: content,
    };
  });
  return {
    chat_history: chat_history,
    message: trimMentionText(slackMsgs[slackMsgs.length - 1].text),
  };
};

/**
 * AIからの応答を取得する関数
 * @param {string} tiggerMsg - ユーザーの入力メッセージ
 * @returns {string} - 応答メッセージ
 */
const fetchAIAnswerText = (tiggerMsg) => {
  // Botに問い合わせられたメッセージを取得する
  const msgsAskedToBot = fetchSlackMsgsAskedToBot(tiggerMsg);
  // Botに問い合わせられたメッセージが無かった場合は空文字を返す
  if (msgsAskedToBot.length === 0) return "";
  // 取得したメッセージをCohere用に変換する
  const msgsForCohere = parseSlackMsgsToCohereQuesryMsgs(msgsAskedToBot);
  // OpenAIへのエンドポイントとAPIキーを設定する
  const ENDPOINT = "https://api.cohere.com/v1/chat";
  // リクエストボディを作成する
  const requestBody = {
    model: model,
    ...msgsForCohere,
    connectors: [{ id: "web-search" }],
  };

  try {
    // CohereへPOSTリクエストを送信する
    const res = UrlFetchApp.fetch(ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: "Bearer " + CO_API_KEY,
        Accept: "application/json",
        "content-type": "application/json",
      },
      contentType: "application/json",
      payload: JSON.stringify(requestBody),
    });

    // ステータスコードが200以外の場合はエラーメッセージを返す
    const resCode = res.getResponseCode();
    if (resCode !== 200) {
      if (resCode === 429) return "利用上限に達しました";
      else return "APIリクエストに失敗しました";
    }

    // レスポンスからAIから返された回答を取得する
    const resPayloadObj = JSON.parse(res.getContentText());

    // 取得した回答を整形して返す
    const rawAnswerText = resPayloadObj.text;
    const trimedAnswerText = rawAnswerText.replace(/^\n+/, "");
    return trimedAnswerText;
  } catch (e) {
    // エラー発生時はエラーメッセージを返す
    console.error(e.stack);
    return `エラーが発生しました ${e.stack}`;
  }
};

const mdToSlack = (raw) => {
  const lines = raw.split("\n");
  result = [];
  lines.forEach((line) => {
    /* Headingとマッチさせる */
    const title = line.match(/^#{1,6} (.+)/);
    if (title) {
      line = `*${title[1]}*`;
    }
    /* Boldとマッチさせる */
    const bold = /\*\*(.+?)\*\*/g;
    line = line.replace(bold, function (match, p1) {
      // p1 は() に一致する部分
      return ` *${p1}* `;
    });
    const boldWithDoubleSpace = / ( \*.+?\* ) /g;
    line = line.replace(boldWithDoubleSpace, function (match, p1) {
      return p1;
    });
    /* 打ち消しとマッチさせる */
    const utikesi = /~~(.+)~~/g;
    line = line.replace(utikesi, function (match, p1) {
      // p1 は() に一致する部分
      return `~${p1}~`;
    });
    /* linkとマッチ */
    const link = /\[([^\]]+)\]\((https:\/\/[^\)]+)\)/g;
    // マッチした部分を置換
    line = line.replace(link, function (match, p1, p2) {
      // p1 はリンクテキスト、p2 はURLに一致
      return `<${p2}|${p1}>`;
    });
    /* Listとマッチ */
    const listRegEx = /^( *)([\*|\-]) (.+)$/;
    // マッチした部分を置換
    line = line.replace(listRegEx, function (match, p1, p2, p3) {
      // p1 はspaceの数、p2は-か*、p3はテキストに一致
      return match.replace(p2, "•  ");
    });
    /* codeBlockStartとマッチ */
    const codeBlockStart = /^```.+$/;
    // マッチした部分を置換
    if (line.match(codeBlockStart)) {
      line = "```";
    }
    result.push(line);
  });
  return result.join("\n");
};

/**
 * Slackの指定したチャンネルにメッセージを投稿する
 * @param {string} channelId - 投稿先のチャンネルID
 * @param {string} message - 投稿するメッセージのテキスト
 * @param {object} option - オプションパラメータ（省略可）
 * @param {string} option.thread_ts - スレッドタイムスタンプ
 * @param {('default'|'primary'|'danger')} option.color - テキスト部分の色
 * @param {boolean} option.link_names - メンションの展開状態を設定
 * @returns {void}
 */
const slackPostMessage = (channelId, message, option) => {
  const url = "https://slack.com/api/chat.postMessage";

  const headers = {
    "Content-Type": "application/json",
    Authorization: "Bearer " + BOT_AUTH_TOKEN,
  };

  // メッセージデータ
  const payload = {
    channel: channelId,
    text: mdToSlack(message),
    ...option,
  };

  // HTTPリクエストを作成
  const options = {
    method: "POST",
    headers,
    payload: JSON.stringify(payload),
  };

  // Slack APIにリクエストを送信
  UrlFetchApp.fetch(url, options);
};

/**
 * doPost関数は、SlackアプリからのPOSTリクエストを処理します。
 * @returns {void}
 */
const doPost = (e) => {
  const reqObj = JSON.parse(e.postData.getDataAsString());

  // Slackから認証コードが送られてきた場合(初回接続時)
  if (reqObj.type == "url_verification") {
    // 認証コードをそのまま返すことで、アプリをSlackに登録する処理が完了する
    return ContentService.createTextOutput(reqObj.challenge);
  }

  // Slackからのコールバック以外の場合、OKを返して処理を終了する
  if (reqObj.type !== "event_callback" || reqObj.event.type !== "message") {
    return ContentService.createTextOutput("OK");
  }

  // メッセージが編集または削除された場合、OKを返して処理を終了する
  if (reqObj.event.subtype !== undefined) {
    return ContentService.createTextOutput("OK");
  }

  // Slackから送信されたトリガーメッセージ
  const triggerMsg = reqObj.event;
  // ユーザーID
  const userId = triggerMsg.user;
  // メッセージID
  const msgId = triggerMsg.client_msg_id;
  // チャンネルID
  const channelId = triggerMsg.channel;
  // タイムスタンプ
  const ts = triggerMsg.ts;

  // Bot自身によるメッセージである場合、OKを返して処理を終了する
  if (userId === BOT_MEMBER_ID) {
    return ContentService.createTextOutput("OK");
  }

  // 処理したメッセージのIDをキャッシュして、同じメッセージを無視する
  const isCachedId = (id) => {
    const cache = CacheService.getScriptCache();
    const isCached = cache.get(id);
    // キャッシュされたIDである場合、trueを返す
    if (isCached) return true;
    // IDをキャッシュに追加する
    cache.put(id, true, 60 * 5); // 5分間キャッシュする
    return false;
  };

  // 処理済みのメッセージの場合、OKを返して処理を終了する
  if (isCachedId(msgId)) {
    return ContentService.createTextOutput("OK");
  }

  try {
    // 応答メッセージを取得する
    const answerMsg = fetchAIAnswerText(triggerMsg);
    // 応答メッセージが存在しない場合、OKを返して処理を終了する
    if (!answerMsg) return ContentService.createTextOutput("OK");
    // Slackに応答メッセージを投稿する
    slackPostMessage(channelId, answerMsg, { thread_ts: ts });
    return ContentService.createTextOutput("OK");
  } catch (e) {
    console.error(e.stack, "応答エラーが発生");
    // エラーが発生した場合、NGを返して処理を終了する
    return ContentService.createTextOutput("NG");
  }
};
