import fs from 'fs';
import path from 'path';
import readline from 'readline';
import dotenv from 'dotenv';
import { AzureOpenAI } from 'openai';
import { fileURLToPath } from 'url';

dotenv.config();

// 入力ファイルのパス
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const FILE_EXCEL_PATH = path.join(__dirname, 'input', 'Excel.zip');
const OUTPUT_FOLDER_PATH = path.join(__dirname, 'output');

// 環境変数から設定を取得
const apiEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
const apiKey = process.env.AZURE_OPENAI_API_KEY;
const apiVersion = process.env.API_VERSION;
const assistantId = process.env.ASSISTANT_ID;
const fileFontId = process.env.FONT_FILE_ID;

///////////////////////////////////
// ファイル アップロード関数
///////////////////////////////////
async function uploadFile(client, filePath, fileType) {
  const fileStream = fs.createReadStream(filePath);
  const file = await client.files.create({
    file: fileStream,
    purpose: 'assistants'
  });
  fileStream.close();
  console.log(`${fileType} file uploaded successfully. File ID: ${file.id}`);
  return file.id;
}

///////////////////////////////////
// ファイル ダウンロード関数
///////////////////////////////////
async function downloadFile(client, fileName, file_id, fileType) {
  try {
    // 元ファイルを取得
    const fileResponse = await client.files.content(file_id);
    const fileData = await fileResponse.buffer();

    // 保存先ディレクトリの確認とファイルのローカル保存
    const outputDir = path.join(OUTPUT_FOLDER_PATH, `${fileType}`);
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }
    const filePath = path.join(outputDir, `${fileName}`);
    fs.writeFileSync(filePath, fileData);
    console.log(`File saved as '${fileType}\\${fileName}'`);

    // 元ファイルを削除
    await client.files.del(file_id)
    console.log("File deleted successfully.")
  } catch (e) {
    console.log(`Error download file: ${e.message}`);
  }
}

///////////////////////////////////
// スレッド作成関数
///////////////////////////////////
async function createThread(client, fileExcelId) {
  const thread = await client.beta.threads.create({
    messages: [{
      role: 'user',
      content: 'アップロードされた Font.zip と Excel.zip を /mnt/data/upload_files に展開してください。これらの ZIP ファイルには解析対象の EXCEL ファイルと日本語フォント NotoSansJP.ttf が含まれています。展開した先にある EXCEL ファイルをユーザーの指示に従い解析してください。EXCEL データからグラフやチャート画像を生成する場合、タイトル、軸項目、凡例等に NotoSansJP.ttf を利用してください。',
      attachments: [
      {
        "file_id": fileFontId,
        "file_id": fileExcelId,
        "tools": [{ "type": "code_interpreter" }]
      }]
    }]
  });
  console.log(`Thread created successfully. Thread ID: ${thread.id}`);
  return thread.id;
}

///////////////////////////////////
// チャットループ関数
///////////////////////////////////
async function chatLoop(client, threadId) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  const question = (query) =>
    new Promise((resolve) => rl.question(query, resolve));

  while (true) {
    // ユーザー入力を取得
    const user_input = await question('\nUser: ');

    // 終了コマンドの処理
    if (user_input.toLowerCase() === 'exit') {
      console.log('Ending session...');
      break;
    }

    // ユーザーのメッセージ送信
    await client.beta.threads.messages.create(
      threadId,
      {
          role: 'user',
          content: user_input,
      }
    );

    // アシスタントの応答を取得
    let run = await client.beta.threads.runs.create(
      threadId,
      {
        assistant_id: assistantId,
      }
    );
    console.log(`Run created:  ${JSON.stringify(run)}`);
    console.log(`\nWaiting for response...`);

    // チャットループ
    while (true) {
      // run の最新状態を取得
      run = await client.beta.threads.runs.retrieve(threadId, run.id);

      // スレッド内の全メッセージを取得
      const messages = await client.beta.threads.messages.list(threadId);

      // 実行状況に応じて処理を分岐
      if (run.status === 'completed') {
        console.log(`\nRun status: ${run.status}`);

        // すべてのメッセージの内容を出力（デバッグトレース用）
        messages.data.forEach(message => {
          console.log(message.content);
        });

        console.log('\nAssistant:');

        // 最初のメッセージの content 配列を処理
        const contentBlocks = messages.data[0].content;
        for (const block of contentBlocks) {
          if (block.type === 'text') {
            let output_text = block.text.value;
            // ファイルパスのアノテーションを処理（後ろから処理することで、インデックスのずれを防ぐ）
            for(let i = block.text.annotations.length - 1; i >= 0; i--) {
              const annotation = block.text.annotations[i];
              if (annotation.type === 'file_path') {
                // ファイル名を取得
                const parts = annotation.text.split('/');
                const fileName = parts[parts.length - 1];
                // ファイルをダウンロード
                await downloadFile(client, fileName, annotation.file_path.file_id, "download_files");
                // value の中の指定された範囲を新しいパスに置換
                const newPath = `output/download_files/${fileName}`;
                output_text = output_text.slice(0, annotation.start_index) + newPath + output_text.slice(annotation.end_index);              
              }
            }
            console.log(output_text);
          } else if (block.type === 'image_file') {
            const fileId = block.image_file.file_id;
            console.log(`[Image file received: ${fileId}]`);
            // 画像ファイルをダウンロード
            const fileName = `${fileId}.png`;
            await downloadFile(client, fileName, fileId, "images");
          } else {
            console.log(`Unhandled content type: ${block.type}`);
          }
        }
        break; // 内部のポーリングループを抜ける

      } else if (run.status === 'queued' || run.status === 'in_progress') {
        console.log(`\nRun status: ${run.status}`);

        // すべてのメッセージの内容を出力（デバッグトレース用）
        messages.data.forEach(message => {
          console.log(message.content);
        });

        // 5秒待機してから再度ポーリング
        await new Promise(resolve => setTimeout(resolve, 5000));

      } else {
        console.log(`Run status: ${run.status}`);
        if (run.status === 'failed') {
          console.log(`Error Code: ${run.last_error.code}, Message: ${run.last_error.message}`);
        }
        break;
      }
    }
  }
  rl.close();
}

///////////////////////////////////
// メイン関数
///////////////////////////////////
async function main() {
  try {
    // 出力フォルダの作成
    if (!fs.existsSync(OUTPUT_FOLDER_PATH)) {
      fs.mkdirSync(OUTPUT_FOLDER_PATH);
    }

    // Azure OpenAI クライアントの初期化
    const client = new AzureOpenAI({
      apiKey: apiKey,
      apiVersion: apiVersion,
      azureEndpoint: apiEndpoint,
    });

    // ファイルのアップロードとスレッドの作成
    const fileExcelId = await uploadFile(client, FILE_EXCEL_PATH, 'Excel');
    const threadId = await createThread(client, fileExcelId);

    // チャットセッションの開始
    console.log("Chat session started. Type 'exit' to end the session.");
    await chatLoop(client, threadId, assistantId);

    // スレッドとファイルの削除
    await client.beta.threads.del(threadId);
    console.log('Thread deleted successfully.');
    await client.files.del(fileExcelId);
    console.log("Excel file deleted successfully.");

  } catch (e) {
    console.error(`An error occurred: ${e}`);
  }
}

main();