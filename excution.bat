@rem .csvファイルのあるフォルダに移動
cd C:\Users\y-nishikawa\Documents\GitHub\Auto_Calc_program


@rem 既に存在しているprocess.pyをrun
python process_final.py

@rem 既に存在しているprocess2.pyをrun
python process_final2.py

python pz31.py 
@REM <for text copy to outlook>

start Outlook	

@REM <Outlook Active)

python pz32.py 
@REM Click F9 key for sending email