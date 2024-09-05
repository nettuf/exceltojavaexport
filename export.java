public static void importExcel(String filePath, Integer idUser) throws IOException, NegocioException {
    CartaoFC cartaoFC = new CartaoFC();
    TopicoFC topicoFC = new TopicoFC();
    TopicoPessoaFC topicoPessoaFC = new TopicoPessoaFC();
    
    FileInputStream fileInputStream = new FileInputStream(filePath);
    Workbook workbook = new XSSFWorkbook(fileInputStream);

    Sheet sheet = workbook.getSheetAt(0);
    int firstDataRow = 1;
    
    for (int i = firstDataRow; i <= sheet.getLastRowNum(); i++) {
        Row row = sheet.getRow(i);
        if (row != null) {
            CartaoFCTO cartaoFCTO = new CartaoFCTO();
            cartaoFCTO.setId(0);
            cartaoFCTO.setSituacao("A");
            cartaoFCTO.setIdCriador(idUser);

            TopicoFCTO topicoFCTO = null;
            Collection<TopicoTemaFCTO> listaTopicoTemaFC = new ArrayList<>();
            Collection<TopicoCursoFCTO> listaTopicoCursoFC = new ArrayList<>();

            for (int j = 0; j < 8; j++) {
                Cell cell = row.getCell(j);
                if (cell != null) {
                    switch (j) {
                        case 0:
                            // Tópico - Nome
                            String topicoNome = getCellValueAsString(cell);
                            if(topicoFC.findByName(topicoNome)!=null) {
                                topicoFCTO = topicoFC.findByName(topicoNome);
                            }else{
                                topicoFCTO = new TopicoFCTO();
                                topicoFCTO.setId(0);
                                topicoFCTO.setNome(topicoNome);
                                topicoFCTO.setIdCriador(idUser);
                                topicoFCTO.setIdProprietario(idUser);
                                topicoFCTO.setSituacao("A");
                                topicoFCTO.setPrivacidade("U");
                                topicoFC.salvar(topicoFCTO);
                                topicoFCTO = topicoFC.findByName(topicoFCTO.getNome());
                            }
                            break;
                        case 1:
                            // Tipo (Texto)
                            String tipo = getCellValueAsString(cell);
                            cartaoFCTO.setTipo(tipo);
                            break;
                        case 2:
                            // Card - Nome
                            String cardNome = getCellValueAsString(cell);
                            cartaoFCTO.setNome(cardNome);
                            break;
                        case 3:
                            // Card - Pergunta
                            String cardPergunta = getCellValueAsString(cell);
                            cartaoFCTO.setPergunta(cardPergunta);
                            break;
                        case 4:
                            // Card - Resposta
                            String cardResposta = getCellValueAsString(cell);
                            cartaoFCTO.setResposta(cardResposta);
                            break;
                        case 5:
                            // tema
                            String idTemas = getCellValueAsString(cell);
                            idTemas = idTemas.trim();
                            String[] listaTema = idTemas.split(",");
                            for (String obj : listaTema) {
                                obj= obj.trim();
                                TopicoTemaFCTO topicoTemaFCTO = new TopicoTemaFCTO();
                                topicoTemaFCTO.setIdTema(Integer.valueOf(obj));
                                topicoTemaFCTO.setIdTopico(topicoFCTO.getId());
                                listaTopicoTemaFC.add(topicoTemaFCTO);
                            }
                            break;
                        case 6:
                            // Tipo + id do Curso (exemplo: P13, T43)
                            String tipoCurso = getCellValueAsString(cell);
                            tipoCurso = tipoCurso.trim();
                            String[] listaTipoCurso = tipoCurso.split(",");
                            for (String obj : listaTipoCurso) {
                                TopicoCursoFCTO topicoCursoFCTO = new TopicoCursoFCTO();
                                topicoCursoFCTO.setId(0);
                                topicoCursoFCTO.setIdTopico(topicoFCTO.getId());
                                
                                obj = obj.trim();
                                String type = obj.substring(0, 1);
                                Integer idType = Integer.valueOf(obj.substring(1));

                                topicoCursoFCTO.setTipo(type);
                                topicoCursoFCTO.setIdCurso(idType);

                                listaTopicoCursoFC.add(topicoCursoFCTO);
                            }
                            break;
                    }
                }
            }
            
            if (topicoFCTO != null) {
                topicoFCTO.setListaTemaFC(listaTopicoTemaFC);
                topicoFCTO.setListaTopicoCursoFC(listaTopicoCursoFC);
                topicoFC.salvar(topicoFCTO);
                topicoFCTO = topicoFC.findByName(topicoFCTO.getNome());

                cartaoFCTO.setIdTopico(topicoFCTO.getId());
                cartaoFC.inserirCartaoFC(cartaoFCTO);

                TopicoPessoaFCTO topicoPessoaFCTO = new TopicoPessoaFCTO();
                topicoPessoaFCTO.setIdPessoa(idUser);
                topicoPessoaFCTO.setIdTopico(topicoFCTO.getId());
                topicoPessoaFCTO.setPerfil("E");
                topicoPessoaFC.inserirTopicoPessoaFC(topicoPessoaFCTO);
            }
        }
    }

    workbook.close();
    fileInputStream.close();
}

public static String getCellValueAsString(Cell cell) {
    if (cell == null) {
        return "";
    }

    switch (cell.getCellType()) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue().toString();
            } else {
                DecimalFormatSymbols symbols = new DecimalFormatSymbols(Locale.getDefault());
                symbols.setDecimalSeparator(',');
                DecimalFormat decimalFormat = new DecimalFormat("##########0.#########", symbols);
                decimalFormat.setGroupingUsed(false); 
                String formattedNumber = decimalFormat.format(cell.getNumericCellValue());
                
                if (formattedNumber.contains(",")) {
                    formattedNumber = formattedNumber.replaceAll("\\.(?=\\d*$)", "");
                }
                
                return formattedNumber;
            }
        case BOOLEAN:
            return Boolean.toString(cell.getBooleanCellValue());
        case FORMULA:
            FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            Cell evaluatedCell = evaluator.evaluateInCell(cell);
            return getCellValueAsString(evaluatedCell);
        default:
            return "";
    }
}