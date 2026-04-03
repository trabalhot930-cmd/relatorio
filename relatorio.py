def gerar_relatorio(doc_path, dados):
    from docx import Document
    from docx.shared import Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    import copy, io

    doc = Document(doc_path)

    def full_text(p):
        return "".join(r.text for r in p.runs)

    def substituir_texto_paragrafo(p, novo_texto):
        """Substitui o texto preservando a formatação do primeiro run."""
        if not p.runs:
            return
        # Copia rPr do primeiro run
        primeiro = p.runs[0]._r
        ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        rPr = primeiro.find(f'{{{ns}}}rPr')
        rPr_copia = copy.deepcopy(rPr) if rPr is not None else None

        # Remove todos os runs
        for run in list(p.runs):
            run._r.getparent().remove(run._r)

        # Cria novo run
        novo_run = p.add_run(novo_texto)
        if rPr_copia is not None:
            novo_run._r.insert(0, rPr_copia)

    def inserir_bloco_texto(p_ref, linhas):
        """Substitui parágrafo {{TAG}} por múltiplas linhas."""
        parent = p_ref._element.getparent()
        idx = list(parent).index(p_ref._element)
        # Remove o parágrafo de referência
        parent.remove(p_ref._element)
        # Insere novos parágrafos
        for i, linha in enumerate(linhas):
            novo_p = OxmlElement('w:p')
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = linha
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            novo_p.append(r)
            parent.insert(idx + i, novo_p)

    # Dados
    numero = dados['numero']
    assunto = dados['assunto']
    data_manut = dados['data']
    local = dados['local']
    texto_informativo = dados['texto_informativo']
    texto_conclusivo = dados['texto_conclusivo']
    nome_foto = dados['nome_foto']
    foto_bytes = dados.get('foto_bytes')
    MESES = dados['MESES']

    data_str = f"{data_manut.day} de {MESES[data_manut.month]} de {data_manut.year}"
    data_upper = data_str.upper()

    for p in doc.paragraphs:
        txt = full_text(p)

        if "RELATÓRIO DE SEGURANÇA" in txt:
            substituir_texto_paragrafo(p, f"RELATÓRIO DE SEGURANÇA Nr. {numero} / 2026")

        elif "DATA:" in txt:
            substituir_texto_paragrafo(p, f"DATA: {data_upper}")

        elif "ASSUNTO:" in txt:
            substituir_texto_paragrafo(p, f"ASSUNTO: {assunto.upper()}")

        elif "{{INFORMATIVA}}" in txt:
            linhas = texto_informativo.split('\n')
            inserir_bloco_texto(p, linhas)

        elif "{{CONCLUSIVA}}" in txt:
            linhas = texto_conclusivo.split('\n')
            inserir_bloco_texto(p, linhas)

        elif "Vitória do Xingu" in txt:
            substituir_texto_paragrafo(p, f"Vitória do Xingu /PA, {data_str}")

    # Foto
    if foto_bytes:
        for i, p in enumerate(doc.paragraphs):
            if "PARTE ILUSTRATIVA" in full_text(p):
                try:
                    doc.paragraphs[i+1].clear()
                    doc.paragraphs[i+1].add_run(nome_foto)
                    doc.paragraphs[i+1].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    doc.paragraphs[i+2].clear()
                    run = doc.paragraphs[i+2].add_run()
                    run.add_picture(io.BytesIO(foto_bytes), width=Cm(16))
                    doc.paragraphs[i+2].alignment = WD_ALIGN_PARAGRAPH.CENTER

                    doc.paragraphs[i+3].clear()
                    doc.paragraphs[i+3].add_run(f"Figura 1 – {nome_foto}")
                    doc.paragraphs[i+3].alignment = WD_ALIGN_PARAGRAPH.CENTER
                except Exception as e:
                    print(f"Erro na imagem: {e}")
                break

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()
