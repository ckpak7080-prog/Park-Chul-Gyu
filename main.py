import win32com.client
import datetime
import os
import time


def download_and_merge_todays_ppts():
    # ==========================================
    # [설정] 부서 및 담당자 리스트 (자료취합순서 포함)
    # ==========================================
    target_list = [
        {"order": 1, "team": "연구기획팀", "staff": "허용민"},
        {"order": 2, "team": "심혈관팀", "staff": "김태원"},
        {"order": 3, "team": "급성감염팀", "staff": "한예지"},
        {"order": 4, "team": "Cancer팀", "staff": "정진용"},
        {"order": 5, "team": "호르몬팀", "staff": "이소희"},
        {"order": 6, "team": "치료용항체팀", "staff": "김영은"},
        {"order": 7, "team": "갑상선팀", "staff": "김세희"},
        {"order": 8, "team": "당뇨팀", "staff": "함은선"}
    ]

    # 1. 파일 저장 경로 설정
    base_dir = os.getcwd()
    download_dir = os.path.join(base_dir, "Meeting_PPTs")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)

    # 오늘 날짜를 YYMMDD 형식으로 변환 (예: 260521)
    today_date = datetime.date.today()
    today_str = today_date.strftime("%y%m%d")

    downloaded_files_with_order = []  # (취합순서, 파일경로) 형태로 저장
    received_teams = set()  # 메일을 제출한 팀을 기록할 셋(Set)

    print(f"[{today_date}] 파일명에 '{today_str}'(이)가 포함된 PPT 첨부파일을 검색합니다...\n")

    # ==========================================
    # 1단계: 오늘 날짜(YYMMDD) PPT 검색 및 담당자 매칭
    # ==========================================
    for msg in messages:
        try:
            if msg.Class != 43:
                continue

            msg_date = datetime.date(msg.ReceivedTime.year, msg.ReceivedTime.month, msg.ReceivedTime.day)

            if msg_date < today_date:
                break
            if msg_date != today_date:
                continue

            attachments = msg.Attachments
            for i in range(1, attachments.Count + 1):
                attachment = attachments.Item(i)
                filename = attachment.FileName

                # 첨부파일이 PPT이고, 파일명에 오늘 날짜가 포함된 경우
                if filename and filename.lower().endswith(('.ppt', '.pptx')):
                    if today_str in filename:
                        time_str = msg.ReceivedTime.strftime("%H%M%S")
                        safe_filename = f"{time_str}_{filename}"
                        save_path = os.path.join(download_dir, safe_filename)

                        # 첨부파일 저장
                        attachment.SaveAsFile(save_path)

                        # -----------------------------------------------------
                        # [담당자 확인 로직 개선] 파일명을 최우선으로 확인
                        # -----------------------------------------------------
                        matched_order = 999
                        matched_team = None

                        # 1순위: 파일명에 부서 이름이 있는지 확인 (가장 정확함)
                        for item in target_list:
                            if item["team"] in filename:
                                matched_order = item["order"]
                                matched_team = item["team"]
                                break

                        # 2순위: 보낸 사람(SenderName) 확인
                        if not matched_team:
                            for item in target_list:
                                if item["staff"] in msg.SenderName:
                                    matched_order = item["order"]
                                    matched_team = item["team"]
                                    break

                        # 3순위: 메일 제목이나 본문에서 팀명/이름 유추
                        if not matched_team:
                            subject_body = (msg.Subject + str(msg.Body)).replace("\n", "")
                            for item in target_list:
                                if item["team"] in subject_body or item["staff"] in subject_body:
                                    matched_order = item["order"]
                                    matched_team = item["team"]
                                    break

                        if matched_team:
                            received_teams.add(matched_team)

                        downloaded_files_with_order.append((matched_order, save_path))
                        print(f"다운로드 완료: {safe_filename} (추정 부서: {matched_team if matched_team else '미상'})")

        except Exception as e:
            continue

    # ==========================================
    # 2단계: 미제출 부서 확인 및 출력
    # ==========================================
    missing_teams = [item for item in target_list if item["team"] not in received_teams]

    print("\n" + "=" * 40)
    print(" 📊 [주간보고서 제출 현황]")
    print("=" * 40)
    print(f" - 수신 완료: {len(received_teams)}팀")
    print(f" - 미제출: {len(missing_teams)}팀")

    if missing_teams:
        print("\n [🚨 미제출 부서 및 담당자 목록]")
        for item in missing_teams:
            print(f"  • {item['order']}순위 | {item['team']} ({item['staff']})")
    else:
        print("\n ✅ 모든 부서가 주간보고서를 제출했습니다!")
    print("=" * 40 + "\n")

    # ==========================================
    # 3단계: 취합순서에 맞게 파일 정렬 후 병합
    # ==========================================
    if not downloaded_files_with_order:
        print(f"파일명에 '{today_str}'이(가) 포함된 PPT 파일이 없습니다. 프로그램을 종료합니다.")
        return

    # 취합순서(order)를 기준으로 오름차순 정렬
    downloaded_files_with_order.sort(key=lambda x: x[0])

    # 정렬된 순서대로 파일 경로만 추출
    files_to_merge = [path for order, path in downloaded_files_with_order]

    print(f"총 {len(files_to_merge)}개의 PPT 파일을 취합 순서에 맞춰 병합합니다...")

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True

    try:
        # 첫 번째 파일을 베이스로 엽니다.
        base_ppt = ppt_app.Presentations.Open(files_to_merge[0])

        # 두 번째 파일부터 차례대로 열어서 복사/붙여넣기 합니다.
        for file_path in files_to_merge[1:]:
            source_ppt = ppt_app.Presentations.Open(file_path)

            source_ppt.Slides.Range().Copy()
            time.sleep(0.5)

            paste_index = base_ppt.Slides.Count
            base_ppt.Slides.Paste(Index=paste_index)

            source_ppt.Close()

        # 병합된 파일을 저장합니다.
        merged_file_name = "주간보고병합.pptx"
        merged_file_path = os.path.join(base_dir, merged_file_name)

        if os.path.exists(merged_file_path):
            os.remove(merged_file_path)

        base_ppt.SaveAs(merged_file_path)
        base_ppt.Close()

        print(f"\n✅ 병합이 완료되었습니다!")
        print(f"저장 위치: {merged_file_path}")

    except Exception as e:
        print(f"\n❌ PPT 병합 중 오류가 발생했습니다: {e}")
    finally:
        ppt_app.Quit()


# 실행
if __name__ == "__main__":
    download_and_merge_todays_ppts()
