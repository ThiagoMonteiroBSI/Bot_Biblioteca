import pandas as pd

df = pd.read_excel('Biblioteca.xlsx', sheet_name='Página1', dtype=str)

col_lidos = df.columns[0] #contagem de colunas
col_sipac = df.columns[1]
col_falta = df.columns[2]

def limpar_dado(valor): #retira dados inconsistentes
    if pd.isna(valor) or str(valor).strip().lower() in ["", "nan"]:
        return ""
    texto = str(valor).strip().replace(" ", "") 
    if texto.endswith(".0"):
        texto = texto[:-2]
    return texto
#Compara os lidos com os sipacs e os faltantes
lidos_originais = [limpar_dado(x) for x in df[col_lidos].tolist()]
sipac_sujos = [limpar_dado(x) for x in df[col_sipac].tolist()]
sipac_originais = [x for x in sipac_sujos if x != ""]


sipac_counts = {}
for s in sipac_originais:
    sipac_counts[s] = sipac_counts.get(s, 0) + 1

new_sipac = []


for l in lidos_originais:
    if l == "":
        new_sipac.append("")
    elif sipac_counts.get(l, 0) > 0:
        new_sipac.append(l)   
        sipac_counts[l] -= 1  
    else:
        new_sipac.append("")  

falta_list = []
for s, count in sipac_counts.items():
    falta_list.extend([s] * count)

max_len = max(len(lidos_originais), len(falta_list))
lidos_final = lidos_originais + [""] * (max_len - len(lidos_originais))
sipac_final = new_sipac + [""] * (max_len - len(new_sipac))
falta_final = falta_list + [""] * (max_len - len(falta_list))

df_resultado = pd.DataFrame({
    col_lidos: lidos_final,
    col_sipac: sipac_final,
    col_falta: falta_final
})

df_resultado.to_excel('Biblioteca_Automatizada.xlsx', index=False)
print(f"Trabalho concluído! {len(falta_list)} exemplares enviados para a coluna FALTA.")