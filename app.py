def preencher_celula_segura(ws, texto_busca, valor, col_offset=1):
    """Preenche células evitando células mescladas e lidando com tuplas"""
    for row in ws.iter_rows():
        for cell in row:
            try:
                # Verifica se o valor da célula é uma tupla
                cell_value = cell.value
                if isinstance(cell_value, tuple):
                    # Se for tupla, pega o primeiro elemento não vazio
                    cell_value = next((str(x) for x in cell_value if x), "")
                
                if cell_value and str(texto_busca).strip() in str(cell_value).strip():
                    try:
                        # Verifica se a célula alvo está mesclada
                        for merged_range in ws.merged_cells.ranges:
                            if (cell.row, cell.column + col_offset) in merged_range:
                                # Preenche a primeira célula do merge
                                ws.cell(row=merged_range.min_row, 
                                       column=merged_range.min_col, 
                                       value=valor)
                                return True
                        
                        # Se não está mesclada, preenche normalmente
                        ws.cell(row=cell.row, column=cell.column+col_offset, value=valor)
                        return True
                    except Exception as e:
                        st.warning(f"Célula {cell.coordinate} não pôde ser preenchida: {str(e)}")
                        return False
            except Exception as e:
                st.warning(f"Erro ao processar célula {cell.coordinate}: {str(e)}")
                continue
    return False