from pathlib import Path

path = Path(r'c:\Users\Usuario\Desktop\py\base\proyecto proveedores\app.py')
text = path.read_text(encoding='utf-8')
text_n = text.replace('\r\n', '\n')
marker = '# ══════════════════════════════════════════════════════════════════════════\n#  HERRAMIENTA: VERIFICAR PRODUCTOS PARA GARANTÍA (Airtable)\n# ══════════════════════════════════════════════════════════════════════════\n'
idx = text_n.find(marker)
print('marker idx', idx)
if idx == -1:
    raise SystemExit('Marker not found')
replacement = '''# ══════════════════════════════════════════════════════════════════════════
#  PANEL DE PRODUCTOS: ÚLTIMO PANEL EN GOOGLE SHEETS
# ══════════════════════════════════════════════════════════════════════════

with st.expander("🔎 PRODUCTOS – Último panel de Google Sheets", expanded=False):

    st.markdown(
        "Este panel carga los datos del último tab disponible en el spreadsheet de Google Sheets "
        "y muestra los productos en una tabla clásica y ordenada."
    )

    col_p1, col_p2 = st.columns([3, 1])
    if col_p2.button("🔄 Actualizar panel", use_container_width=True, key="panel_refresh"):
        leer_ultimo_panel.clear()
        st.experimental_rerun()

    panel_title, df_panel = leer_ultimo_panel(SHEET_ID)
    if df_panel.empty:
        st.info(
            f"No se encontraron registros en el último panel de Google Sheets "
            f"({panel_title})."
        )
    else:
        st.caption(f"Pestaña leída: **{panel_title}** | Filas: **{len(df_panel)}**")
        st.dataframe(df_panel, use_container_width=True, hide_index=True)

        buf_panel = io.BytesIO()
        with pd.ExcelWriter(buf_panel, engine="xlsxwriter") as writer:
            df_panel.to_excel(writer, index=False, sheet_name=panel_title[:31])
        buf_panel.seek(0)

        col_p1.download_button(
            label="📥 Exportar último panel a Excel",
            data=buf_panel.getvalue(),
            file_name=f"Panel_{panel_title.replace(' ', '_')}_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
'''
text_new = text_n[:idx] + replacement
path.write_text(text_new.replace('\n', '\r\n'), encoding='utf-8')
print('done')
