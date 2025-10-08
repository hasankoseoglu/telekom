import pandas as pd

# Giriş dosyası adı
input_file = "Abone rehber.abn.xlsx"

# Çıkış dosyası adı
output_file = "Abone_rehber_duzenli.xlsx"

# Veriyi oku (tek sütun olarak, metin olarak işle)
df = pd.read_excel(input_file, header=None, dtype=str)

# ;| işaretine göre böl
df_split = df[0].str.split(r";\|", expand=True)

# Yeni dosyaya yaz
df_split.to_excel(output_file, index=False, header=False)

print("✅ Dosya başarıyla dönüştürüldü:", output_file)
