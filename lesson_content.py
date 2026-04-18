"""
主題內容模板 — 根據輸入主題產出完整的 content dict
用於 pptx_generator.process_slides()

優先使用 Gemini API 產出客製化內容，失敗時自動 fallback 到靜態模板。
"""

import json
import os

from dotenv import load_dotenv

load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# 範本 TOC 的原始 key — pptx_generator 用這些 key 做文字比對替換
_TOC_OLD_KEYS = [
    "1.學習目標",
    "2.相撲是什麼",
    "3.相撲的歷史",
    "4.場地介紹",
    "5.比賽規則",
    "6.材料準備",
    "7.動手時間",
    "8.程式範例",
    "9.創意改造",
    "10.課堂作業",
]


def _build_prompt(topic: str) -> str:
    """Build Gemini prompt that returns a JSON dict matching pptx_generator schema."""
    toc_keys_json = json.dumps(_TOC_OLD_KEYS, ensure_ascii=False)

    return f"""你是 AIRDA（人工智慧與機器人發展協會）的資深教案設計專家。
請根據主題「{topic}」產出一份完整的 Matrix Mini R4 機器人教案內容。
適用對象：國小高年級到國中。語言：繁體中文。

請回傳一個 JSON 物件，包含以下欄位：

1. "cover_title": 字串，教案標題（就是主題名稱）
2. "toc": 一個 dict，key 是以下固定的舊標題 {toc_keys_json}，value 是對應的新標題（根據主題改寫，保持編號格式如 "1.學習目標"）
3. "objectives_knowledge": 字串，知識學習目標（2-3 點，用換行分隔，第一行是 "知識學習"）
4. "objectives_practice": 字串，動手實作目標（2-3 點，用換行分隔，第一行是 "動手實作"）
5. "theory_1_title": 字串，理論第一頁標題（如「{topic}是什麼?」）
6. "theory_1_body": 字串，理論第一頁內容（用淺顯的比喻解釋概念，5-8 行，用換行分隔）
7. "theory_2_title": 字串，理論第二頁標題（如「{topic}的原理」）
8. "theory_2_body": 字串，理論第二頁內容（感測原理、訊號轉換、關鍵參數，8-10 行）
9. "theory_3_title": 字串，理論第三頁標題（如「{topic}的應用」）
10. "theory_3_body": 字串，理論第三頁主要內容（3 行概述）
11. "theory_3_sub1": 字串，應用子類別 1（如「工業應用」）
12. "theory_3_sub2": 字串，應用子類別 2（如「生活應用」）
13. "theory_3_sub3": 字串，應用子類別 3（如「競賽應用」）
14. "theory_4_title": 字串，理論第四頁標題（如「生活中的{topic}」）
15. "theory_4_body": 字串，理論第四頁內容（具體生活實例，6-8 行）
16. "theory_4_link": 字串，可留空 ""
17. "transition": 字串，過渡頁文字（鼓勵學生準備動手，3 行）
18. "hands_on": 字串，接線步驟（6 步驟，針對 Matrix Mini R4 的 I2C 接口）
19. "arduino_code": 字串，Arduino C++ 程式碼範例（使用 MatrixMiniR4.h 函式庫，包含 setup 和 loop，有具體的感測器讀取邏輯，不只是註解）
20. "creative_title": 字串，創意挑戰標題
21. "creative_item1": 字串，挑戰玩法說明
22. "creative_item2": 字串，得分規則
23. "creative_item3": 字串，犯規規則
24. "teacher_notes": 字串，教師備註（4 點建議，用換行分隔）
25. "extension": 字串，延伸探討（3 個思考問題，用換行分隔）

重要規則：
- 所有字串使用 \\n 作為換行
- arduino_code 必須是可編譯的 C++ 程式碼
- toc 的 key 必須完全匹配我給你的舊標題列表
- 回傳純 JSON，不要包含 markdown code fence"""


def _generate_with_gemini(topic: str) -> dict:
    """Call Gemini API and parse the JSON response."""
    from google import genai

    client = genai.Client(api_key=GEMINI_API_KEY)

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=_build_prompt(topic),
        config={
            "response_mime_type": "application/json",
            "temperature": 0.7,
        },
    )
    return json.loads(response.text)


def _static_fallback(topic: str) -> dict:
    """靜態模板 fallback — 當 Gemini API 不可用時使用"""

    return {
        'cover_title': f'{topic}',
        'toc': {
            '1.學習目標': '1.學習目標',
            '2.相撲是什麼': f'2.{topic}是什麼?',
            '3.相撲的歷史': f'3.{topic}的原理',
            '4.場地介紹': f'4.{topic}的應用',
            '5.比賽規則': f'5.生活中的{topic}',
            '6.材料準備': '6.材料準備',
            '7.動手時間': '7.動手時間',
            '8.程式範例': '8.程式範例',
            '9.創意改造': f'9.{topic}挑戰賽',
            '10.課堂作業': '10.課堂作業',
        },
        'objectives_knowledge': (
            f'知識學習\n'
            f'1.了解{topic}的基本原理與科學背景\n'
            f'2.認識{topic}在生活中的應用場景'
        ),
        'objectives_practice': (
            f'動手實作\n'
            f'1.組裝{topic}感測器並正確接線\n'
            f'2.寫程式讀取感測器數值並分析數據'
        ),
        'theory_1_title': f'{topic}是什麼?',
        'theory_1_body': (
            f'你有沒有想過，機器人怎麼感知周圍環境？\n'
            f'\n'
            f'就像人有眼睛、耳朵、皮膚來感覺世界，\n'
            f'機器人也需要各種感測器來「感覺」。\n'
            f'\n'
            f'{topic}就是機器人的一種「感覺器官」，\n'
            f'它能把物理世界的訊號轉換成數字，\n'
            f'讓程式可以讀取和判斷。\n'
            f'\n'
            f'今天我們就來認識這個厲害的小元件！'
        ),
        'theory_2_title': f'{topic}的原理',
        'theory_2_body': (
            f'{topic}的核心原理：\n'
            f'\n'
            f'感測器偵測物理量（如光、聲音、距離、溫度）\n'
            f'→ 內部晶片將類比訊號轉換成數位資料\n'
            f'→ 透過 I2C 或 UART 通訊傳給主控板\n'
            f'→ 程式讀取數值進行判斷\n'
            f'\n'
            f'重要參數：\n'
            f'• 量測範圍：感測器能偵測的最大/最小值\n'
            f'• 精確度：量測值與真實值的誤差\n'
            f'• 響應時間：從偵測到回傳結果的速度'
        ),
        'theory_3_title': f'{topic}的應用',
        'theory_3_body': (
            f'{topic}在各種領域都有重要應用：\n'
            f'工業自動化、智慧家居、機器人競賽...\n'
            f'今天學的技術就是這些應用的基礎！'
        ),
        'theory_3_sub1': '工業應用',
        'theory_3_sub2': '生活應用',
        'theory_3_sub3': '競賽應用',
        'theory_4_title': f'生活中的{topic}',
        'theory_4_body': (
            f'你可能沒注意到，{topic}就在你身邊：\n'
            f'\n'
            f'智慧手機裡有各種感測器協助運作\n'
            f'掃地機器人用感測器偵測障礙物\n'
            f'自動門感應到人就會打開\n'
            f'汽車的安全系統時刻偵測周圍環境\n'
            f'\n'
            f'今天學的感測器就是這些科技的縮小版！'
        ),
        'theory_4_link': '',
        'transition': (
            f'同學們\n'
            f'準備好變成\n'
            f'「{topic}達人」了嗎？'
        ),
        'hands_on': (
            f'接線步驟：\n'
            f'1. 確認電池已關閉電源\n'
            f'2. 將{topic}感測器的 I2C 線\n'
            f'   接到 Mini R4 的 I2C1 接口\n'
            f'3. 裝入 2-cell 18650 電池\n'
            f'4. USB 線連接電腦與 Mini R4\n'
            f'5. 打開 Arduino IDE\n'
            f'   選擇正確的板子與 COM Port\n'
            f'6. 上傳程式，打開 Serial Monitor'
        ),
        'arduino_code': (
            '#include <MatrixMiniR4.h>\n'
            '\n'
            'void setup() {\n'
            '  Serial.begin(115200);\n'
            '  MiniR4.begin();\n'
            f'  // 初始化{topic}感測器\n'
            '  // MiniR4.I2C1.Sensor.begin();\n'
            f'  Serial.println("{topic} Ready!");\n'
            '}\n'
            '\n'
            'void loop() {\n'
            '  // 讀取感測器數值\n'
            '  // int value = MiniR4.I2C1.Sensor.getValue();\n'
            '  // Serial.print("Value: ");\n'
            '  // Serial.println(value);\n'
            '  delay(500);\n'
            '}'
        ),
        'creative_title': f'{topic}挑戰賽',
        'creative_item1': f'玩法：用{topic}完成指定任務',
        'creative_item2': '得分：完成基礎得1分 進階得3分',
        'creative_item3': '犯規：偷看別組答案=0分',
        'teacher_notes': (
            f'1.挑戰賽玩法：老師設定目標，學生用'
            f'{topic}感測器完成量測任務\n'
            f'\n'
            f'2.進階：嘗試不同環境條件下的量測，'
            f'引導學生觀察數據變化並思考原因\n'
            f'\n'
            f'3.幽默引導：「感測器說的跟你猜的'
            f'一樣嗎？讓我們來看看誰比較準！」\n'
            f'\n'
            f'4.拍照分享給家長，交流心得'
        ),
        'extension': (
            f'1.{topic}感測器的量測極限在哪裡？\n'
            f'  (提示：想想規格書上的範圍)\n'
            f'\n'
            f'2.在不同環境下量測結果會不同嗎？\n'
            f'  為什麼？\n'
            f'\n'
            f'3.如果要讓機器人自動完成一個任務，\n'
            f'  你會怎麼用{topic}感測器？'
        ),
    }


def generate_content(topic: str) -> dict:
    """主入口：優先 Gemini API，失敗自動 fallback 靜態模板。"""
    if GEMINI_API_KEY:
        try:
            result = _generate_with_gemini(topic)
            print(f"[INFO] Gemini API 成功產出「{topic}」教案內容")
            return result
        except Exception as e:
            print(f"[WARN] Gemini API 失敗，使用靜態模板: {e}")
    return _static_fallback(topic)
