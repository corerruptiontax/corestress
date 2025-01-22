# Corestress

[Saya sedang malas merapihkan, nanti saya update, cara pakainya]

Format source sebelum run move.py
FP_NamaBulan_Tahun.xlsx

di A1 = =IF(COUNTIFS(B$2:B2, B2, C$2:C2, C2, E$2:E2, E2, D$2:D2,D2 )=1, MAX(A$1:A1)+1, A1)
![image](https://github.com/user-attachments/assets/719de225-2910-478b-b90d-a9b11bde0448)


python move.py "sumber_file.xlsx" "file_tujuan/output_setup.xlsx"
