@rem "mvno_pj_202106_calc"ファルダに移動
cd C:\Users\xxx_xxxx\Documents\GitHub\Auto_Calc_program

python process_final.py

python process_final2.py

python pz31.py 
@REM <for text copy to outlook>

start Outlook
@REM <Outlook Active)

python pz32.py 
@REM Click F9 key for sending email