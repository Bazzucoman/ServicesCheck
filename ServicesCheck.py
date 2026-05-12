import os
import pandas as pd

Homologacion = r"C:\Users\JuanM.RamirezM\Desktop\JSON\SURA PIP\TABLAS\Homologacion.xlsx"

consolidado_final = r"C:\Users\JuanM.RamirezM\Documents\SURA PIP\MAYO 2026\PIP CA PULMON PREVALENTE\PIP CA PULMON PREVALENTE.xlsx"

ArchivoProceso = r"C:\Users\JuanM.RamirezM\Documents\SURA PIP\MAYO 2026\PIP CA PULMON PREVALENTE\MEDIO\ArchivoProceso.xlsx"

RutaFinal = r"C:\Users\JuanM.RamirezM\Documents\SURA PIP\MAYO 2026\PIP CA PULMON PREVALENTE"

def approval():

    #REALIZAR UN DICCIONARIO CON LOS CODIGOS HOMOLOGADOS
    dfhm = pd.read_excel(Homologacion, dtype=str)

    Scups = dict(zip(dfhm["SURA_CUPS"], dfhm["CODIGO_ATC"]))

    #LEER EL DOCUMENTO DEL MEDIO MAGNETICO

    dfmt = pd.read_excel(ArchivoProceso, dtype=str)

    dfmt["CUPS"] = dfmt["CUPS"].apply(lambda cups: Scups.get(cups, cups))

    dfmt["NOMBRE_MEDIO"] = dfmt["NOMBRES"].str.split().str[0].str.upper()

    return dfmt

def cruce_auto(df_medio):

    df_relacion = pd.read_excel(consolidado_final, dtype=str)

    
    df_relacion.loc[
        df_relacion["CUPS/CUM"].astype(str).str.contains("-", na=False),
        "CRUCE_NOMBRES"
    ] = (
        df_relacion["Descripción"]
        .str.split()
        .str[0]
        .str.upper()
    )


    df_relacion = df_relacion.merge(
        df_medio[["NUMERO IDENTIFICACION AFILIADO","CUPS","CONSEC_AUT","ESTADO"]],
        left_on=["Número ID","CUPS/CUM"],
        right_on=["NUMERO IDENTIFICACION AFILIADO","CUPS"],
        how="left"
    )

    df_relacion = df_relacion.drop(columns=["NUMERO IDENTIFICACION AFILIADO","CUPS"])

    df_relacion = df_relacion.rename(columns={"CONSEC_AUT": "Autorización"})
    
    df_relacion = df_relacion.rename(columns={"ESTADO": "ESTADO_FINAL"})

    df_relacion = df_relacion.merge(
    df_medio[["NUMERO IDENTIFICACION AFILIADO","NOMBRE_MEDIO","CONSEC_AUT","ESTADO"]],
    left_on=["Número ID","CRUCE_NOMBRES"],
    right_on=["NUMERO IDENTIFICACION AFILIADO","NOMBRE_MEDIO"],
    how="left"
)

    df_relacion["Autorización"] = df_relacion["Autorización"].fillna(df_relacion["CONSEC_AUT"])

    df_relacion["ESTADO_FINAL"] = df_relacion["ESTADO_FINAL"].fillna(df_relacion["ESTADO"])



    df_relacion = df_relacion.drop(columns=["NUMERO IDENTIFICACION AFILIADO", "NOMBRE_MEDIO", "CRUCE_NOMBRES", "CONSEC_AUT", "ESTADO"])


    ruta_salida = os.path.join(RutaFinal, "ServicesCheck.xlsx")

    df_relacion.to_excel(ruta_salida, index=False)

    print("Archivo exportado correctamente")





df_medio = approval()
cruce_auto(df_medio)
