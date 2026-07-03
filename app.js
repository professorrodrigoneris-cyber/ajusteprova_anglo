document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileList = document.getElementById('file-list');
    const btnProcessar = document.getElementById('btn-processar');
    
    const chkNotas = document.getElementById('chk-notas');
    const chkEstilos = document.getElementById('chk-estilos');
    const notasConfig = document.getElementById('notas-config');
    const optionCards = document.querySelectorAll('.option-card');

    let selectedFiles = [];

    // --- Interações da UI ---

    // Toggle states
    chkNotas.addEventListener('change', (e) => {
        optionCards[0].classList.toggle('disabled', !e.target.checked);
    });

    chkEstilos.addEventListener('change', (e) => {
        optionCards[1].classList.toggle('disabled', !e.target.checked);
    });

    // Drag and Drop
    dropZone.addEventListener('click', () => fileInput.click());

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        let dt = e.dataTransfer;
        let files = dt.files;
        handleFiles(files);
    });

    fileInput.addEventListener('change', function() {
        handleFiles(this.files);
    });

    function handleFiles(files) {
        const newFiles = Array.from(files).filter(file => file.name.endsWith('.xlsx') && !file.name.startsWith('~$'));
        selectedFiles = [...selectedFiles, ...newFiles];
        updateFileList();
        btnProcessar.disabled = selectedFiles.length === 0;
    }

    function updateFileList() {
        fileList.innerHTML = '';
        selectedFiles.forEach((file, index) => {
            const item = document.createElement('div');
            item.className = 'file-item';
            item.innerHTML = `
                <span>📄 ${file.name}</span>
                <button style="background:none;border:none;color:#ef4444;cursor:pointer;" onclick="removeFile(${index})">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line></svg>
                </button>
            `;
            fileList.appendChild(item);
        });
    }

    window.removeFile = function(index) {
        selectedFiles.splice(index, 1);
        updateFileList();
        btnProcessar.disabled = selectedFiles.length === 0;
    }

    // --- Processamento ---

    btnProcessar.addEventListener('click', async () => {
        const doNotas = chkNotas.checked;
        const doEstilos = chkEstilos.checked;
        
        if (!doNotas && !doEstilos) {
            alert('Selecione pelo menos uma operação (Ajustar Notas ou Ajustar Estilo).');
            return;
        }

        const notaMaxAlvo = parseInt(document.getElementById('nota-max-alvo').value);
        const notaMinNova = parseInt(document.getElementById('nota-min-nova').value);
        const notaMaxNova = parseInt(document.getElementById('nota-max-nova').value);

        const progressContainer = document.getElementById('progress-container');
        const progressFill = document.getElementById('progress-fill');
        const progressText = document.getElementById('progress-text');

        btnProcessar.disabled = true;
        progressContainer.classList.remove('hidden');

        try {
            for (let i = 0; i < selectedFiles.length; i++) {
                const file = selectedFiles[i];
                progressText.innerText = `Processando: ${file.name} (${i + 1}/${selectedFiles.length})`;
                progressFill.style.width = `${((i) / selectedFiles.length) * 100}%`;
                
                // Ler arquivo usando File Reader e ExcelJS
                const arrayBuffer = await file.arrayBuffer();
                const workbook = new ExcelJS.Workbook();
                await workbook.xlsx.load(arrayBuffer);

                // Base nome da turma
                let nomeTurma = file.name.replace('.xlsx', '');
                if (nomeTurma.endsWith('_2')) nomeTurma = nomeTurma.slice(0, -2); // caso existam sufixos _2

                workbook.eachSheet((worksheet, sheetId) => {
                    // --- AJUSTE DE NOTAS ---
                    if (doNotas) {
                        worksheet.eachRow((row, rowNumber) => {
                            if (rowNumber === 1) return; // Pula a primeira linha (possível cabeçalho)
                            
                            row.eachCell((cell, colNumber) => {
                                const val = cell.value;
                                if (val !== null && val !== undefined && typeof val === 'number') {
                                    if (val <= notaMaxAlvo) {
                                        // Random entre notaMinNova e notaMaxNova
                                        const nova = Math.floor(Math.random() * (notaMaxNova - notaMinNova + 1)) + notaMinNova;
                                        cell.value = nova;
                                    }
                                }
                            });
                        });
                    }

                    // --- AJUSTE DE ESTILO ---
                    if (doEstilos) {
                        const maxCol = Math.min(worksheet.columnCount, 6); // Limite F

                        // Verifica se já tem cabeçalho
                        let jaTemCabecalho = false;
                        const a1Val = worksheet.getCell('A1').value;
                        if (a1Val && (a1Val.toString() === nomeTurma || a1Val.toString().toLowerCase().includes('ano'))) {
                            if (worksheet.getCell('A2').value === 'Nome') {
                                jaTemCabecalho = true;
                            }
                        }

                        if (!jaTemCabecalho) {
                            worksheet.spliceRows(1, 0, []); // Insere linha em branco no topo
                            const a1 = worksheet.getCell('A1');
                            a1.value = nomeTurma;
                            worksheet.mergeCells('A1:F1');
                            a1.font = { bold: true };
                            a1.alignment = { horizontal: 'center' };
                        }

                        const rowCount = worksheet.rowCount;
                        
                        // Preparar estilos reutilizáveis
                        const borderHair = { style: 'hair', color: { argb: 'FF000000' } };
                        const alignCenter = { horizontal: 'center', vertical: 'middle' };
                        const fillGray = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

                        for (let rowIdx = 2; rowIdx <= rowCount; rowIdx++) {
                            const row = worksheet.getRow(rowIdx);
                            const isHeader = (row.getCell(1).value === "Nome");

                            for (let colIdx = 1; colIdx <= maxCol; colIdx++) {
                                const cell = row.getCell(colIdx);
                                
                                // Bordas
                                const currentBorder = cell.border || {};
                                cell.border = {
                                    top: currentBorder.top,
                                    left: currentBorder.left,
                                    right: currentBorder.right,
                                    diagonal: currentBorder.diagonal,
                                    bottom: borderHair
                                };

                                // Alinhamento
                                if (colIdx >= 2) {
                                    cell.alignment = alignCenter;
                                }

                                // Fundo cinza na linha de nomes/matérias
                                if (isHeader) {
                                    cell.fill = fillGray;
                                }
                            }
                        }

                        // Auto-ajuste de colunas
                        for (let colIdx = 1; colIdx <= worksheet.columnCount; colIdx++) {
                            let maxLen = 0;
                            const column = worksheet.getColumn(colIdx);
                            
                            // Iteramos os valores
                            column.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
                                if (rowNumber > 1) { // ignora a linha 1 mesclada
                                    const text = cell.value ? cell.value.toString() : '';
                                    if (text.length > maxLen) {
                                        maxLen = text.length;
                                    }
                                }
                            });

                            column.width = (maxLen + 2) * 1.15;
                        }
                    }
                });

                // Gerar e Salvar Arquivo
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
                
                let outName = file.name;
                if (!outName.includes('_ajustado')) {
                    outName = outName.replace('.xlsx', '_ajustado.xlsx');
                }
                
                saveAs(blob, outName);
                
                progressFill.style.width = `${((i + 1) / selectedFiles.length) * 100}%`;
                
                // Pequeno delay para a UI atualizar
                await new Promise(r => setTimeout(r, 200));
            }

            progressText.innerText = "✅ Processamento Concluído!";
            progressText.style.color = "var(--success)";
            
        } catch (error) {
            console.error(error);
            progressText.innerText = "❌ Erro no processamento: " + error.message;
            progressText.style.color = "#ef4444";
        } finally {
            btnProcessar.disabled = false;
        }
    });
});
