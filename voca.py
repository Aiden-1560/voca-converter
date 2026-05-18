def process_images_safely(client, uploaded_files, api_key, progress_bar, status_text):
    all_data = []
    total_files = len(uploaded_files)
    
    prompt = "영어 단어 추출 프롬프트 내용 (기존 파일 설정 유지)"
    
    current_displayed_percent = 0
    
    for idx, file in enumerate(uploaded_files):
        target_percent = int(((idx + 1) / total_files) * 100)
        
        # 1. 대기 및 처리 중 안내 (상단 모래시계)
        while current_displayed_percent < target_percent - 3:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            # ⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. 문구 노출
            status_text.markdown(f"⏳ 처리하는데 시간이 걸리니 조금만 기다려주세요.. ({total_files - idx}초 남음)")
            time.sleep(0.01)
            
        try:
            file.seek(0)
            image_bytes = file.read()
            
            response = client.models.generate_content(
                model='gemini-2.5-flash',
                contents=[types.Part.from_bytes(data=image_bytes, mime_type=file.type), prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            page_data = json.loads(response.text)
            all_data.extend(page_data)
            
        except Exception as e:
            # ⚠️ 구글 무료 제한(ServerError 또는 RESOURCE_EXHAUSTED) 에러 발생 시 처리
            error_msg = str(e).upper()
            
            # 사진을 여러 장 올렸는데 중간에 끊긴 경우 (이미 몇 장은 성공한 상태)
            if idx > 0:
                st.warning("⚠️ 구글 계정의 하루 무료 사용량(20장)이 모두 마감되었습니다. 프로그램 보호를 위해 현재까지 변환된 파일들로만 워드를 생성합니다.")
                break # 루프를 탈출하여 현재까지 쌓인 데이터(`all_data`)로 워드 생성을 진행시킵니다.
                
            # 첫 번째 장부터 바로 실패한 경우 (시작조차 할 수 없는 상태)
            else:
                st.error("❌ 오늘 사용 가능한 구글 무료 제공량(20장)을 모두 초과하여 변환을 시작할 수 없습니다. 내일 다시 시도해 주세요.")
                return None # 즉시 중단
                
        while current_displayed_percent < target_percent:
            current_displayed_percent += 1
            progress_bar.progress(current_displayed_percent / 100)
            time.sleep(0.01)
            
        time.sleep(0.2)
        
    if all_data:
        progress_bar.progress(1.0)
        status_text.success("🌿 정제 프로세스가 성공적으로 완료되었습니다. (100%)")
    return all_data
