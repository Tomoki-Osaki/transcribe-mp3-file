### pip install pydub SpeechRecognition python-docx tqdm ###
import time, datetime, math, os, sys, tqdm
import speech_recognition as sr 
import pydub, docx

Path = str
# Global: base_file_name, wav_file, wav_parts, 

def input_file_loop() -> Path:
    while True:
        print('\n例のようにmp3ファイルのパスを入力してください。例： C:/Users/someone/Downloads/recording.mp3')
        original_path = input('mp3ファイルのパスをここに入力してください。: ')
        if original_path == '':
            print('プログラムを終了します。')
            sys.exit()
        elif os.path.isfile(original_path) == False: 
            print('\nファイルが見つかりません。パスをダブルコーテーションマークで囲んだり、スラッシュ・バックスラッシュの向きを変えてみてください。')
            continue
        break
    
    return original_path
        
def make_path_no_extention(path: Path) -> None:
    global base_file_name
    
    path_str_list = path.split('/')
    base_file_name = path_str_list[-1]
    base_file_name = base_file_name.split('.')[0]
    common_path = ''
    for dir_name in path_str_list[:-1]:
        common_path += dir_name + '/'
        
    base_file_name = common_path + base_file_name # global
    
    return None 

def mp3_to_wav(mp3_file: Path, base_file_name: Path) -> pydub.AudioSegment.from_wav:
    global wav_file, wav_parts
    
    wav_file = base_file_name + '.wav' # global
    wav_parts = base_file_name + '_parts.wav' # global
    
    recording_file = pydub.AudioSegment.from_mp3(mp3_file)
    recording_file.export(base_file_name + '.wav', format='wav')
    recording_data = pydub.AudioSegment.from_wav(wav_file)

    return recording_data

def transcribing_process(recording_data: pydub.AudioSegment.from_wav,
                         word_file_path: Path) -> None:
    
    file_length_ms = math.ceil(recording_data.duration_seconds * 1000)
    first_60s = [i for i in range(0, file_length_ms-60*1000, 60*1000)]
    last_60s = [i + 60*1000 for i in first_60s]
    
    result_word_file = docx.Document() 
    result_word_file.save(word_file_path) 
    
    recognizer = sr.Recognizer()
    
    for x, y in tqdm.tqdm(zip(first_60s, last_60s), total=len(first_60s)):
        parts = recording_data[x:y+1*300]
        file = parts.export(wav_parts, format='wav')
        with sr.AudioFile(file) as source:
            audio = recognizer.record(source)
            text = recognizer.recognize_google(audio, language='ja-JP')
            result_word_file.add_paragraph(text) 
            result_word_file.save(word_file_path) 
        
    last = recording_data[last_60s[-1]:]
    file = last.export(wav_parts, format='wav')
    with sr.AudioFile(file) as source:
        audio = recognizer.record(source)
    text = recognizer.recognize_google(audio, language='ja-JP')
    result_word_file.add_paragraph(text)
    result_word_file.save(word_file_path)

    return None

def delete_unnecessary_files(files_to_delete: list[Path]) -> None:
    for path in files_to_delete:
        if os.path.isfile(path) == True: 
            os.remove(path)

def main():
    mp3_file = input_file_loop()
    
    make_path_no_extention(mp3_file) # global: base_file_name
    
    recording_data = mp3_to_wav(mp3_file, base_file_name)
    
    start_time = datetime.datetime.now().strftime('%H:%M:%S') 
    start = time.time() 
    print(f'\n実行開始時刻は{start_time}です。\n')
    
    file_length_min = math.ceil((recording_data.duration_seconds) / 60)
    
    print(f'\n\nプログラムの実行が終了するまで、おおよそ音声ファイルの長さの半分の時間が掛かります。\n\
今回の音声ファイルは約{file_length_min}分なので、終了まで約{file_length_min/2}分かかります。\n\
実行時間は、実行中にパソコンで行う他のタスク（ネット検索やWord、Excelなどの他のソフトの使用など）の量によって前後します。\n\
ただし、動画を視聴すると処理が重くなってプログラムが中断される可能性があるため、実行中の動画の視聴は念のためお控えください。\n\
必ずインターネットに接続して、パソコンを充電しながら実行してください。\n\
それではプラグラムの実行が終了するまで、しばらくお待ちください。\n\n\
---------------------------------- 実行中 ------------------------------------------\n\n')
    
    word_file_path = base_file_name + '.docx'
    transcribing_process(recording_data, word_file_path)
    
    end = time.time()
    duration = (end - start) / 60
    end_time = datetime.datetime.now().strftime('%H:%M:%S') 
    print(f'\n実行終了時刻は{end_time}です。実行時間は{duration:.2f}分でした。\n\
作成されたWordファイルは {word_file_path} にあります。\n')
        
    
if __name__ == '__main__':
    print('作成者：大崎智貴\nEmail: ootmootmk@gmail.com')
    while True:
        main()
         
        # main中だとファイルが消せないため、delete_unnecessary_filesはmainの終了後に実行する
        delete_unnecessary_files([wav_file, wav_parts])
    
        go_next = input('\n\n続けて別のファイルの書き起こしをしますか？(y/n): ')
        if go_next == 'y': 
            continue
        else: 
            break
