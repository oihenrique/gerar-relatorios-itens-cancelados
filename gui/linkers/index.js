async function create_spreadsheet_from_clipboard() {
    try {
        const dateInput = document.getElementById('date-input').value;
    
        if (!dateInput) {
            throw new Error('verifique a data inserida e tente novamente.');
        }
    
        try {
            let options = {
                mode: 'text',
                pythonPath: window.python.path('../.venv/Scripts/python.exe'),
                args: ['create_spreadsheet_from_clipboard', formatDate(dateInput)]
            }
            const scriptPath = window.python.path('../main.py');
    
            const result = await window.python.runPython(scriptPath, options);
            logMessage('Planilha geral criada.');
        } catch (error) {
            logError('Erro durante a execução do script Python:', error);
        }

    } catch (error) {
        logError('Data inválida:', error);
    }
}

async function generate_spreadsheets() {
    try {
        const dateInput = document.getElementById('date-input').value;
    
        if (!dateInput) {
            throw new Error('verifique a data inserida e tente novamente.');
        }

        try {
            let options = {
                mode: 'text',
                pythonPath: window.python.path('../.venv/Scripts/python.exe'),
                args: ['generate_spreadsheets', formatDate(dateInput)]
            }
            const scriptPath = window.python.path('../main.py');

            const result = await window.python.runPython(scriptPath, options);
            logMessage(`As planilhas de cada loja foram criadas.`);
        } catch (error) {
            logError('Erro durante a execução do script Python:', error);
        }

    } catch (error) {
        logError('Data inválida:', error);
    }
}

async function send_email() {
    try {
        const dateInput = document.getElementById('date-input').value;
    
        if (!dateInput) {
            throw new Error('verifique a data inserida e tente novamente.');
        }

        try {
            let options = {
                mode: 'text',
                pythonOptions: ['-u'],
                pythonPath: window.python.path('../.venv/Scripts/python.exe'),
                args: ['send_email', formatDate(dateInput)]
            }
            const scriptPath = window.python.path('../main.py');

            const result = await window.python.runPython(scriptPath, options);
            logMessage('Lembre-se de abrir o aplicativo de e-mails antes de prosseguir.')
            logMessage('Envio em andamento. Verifique a caixa de entrada pelo aplicativo no computador para acompanhar o progresso.')
            logMessage('Não feche este programa até que todos sejam enviados.')

        } catch (error) {
            logError('Erro durante a execução do script Python:', error);
        }
    } catch (error) {
        logError('Data inválida:', error);
    }
}

function formatDate(inputDate) {
    const parts = inputDate.split('-');
    if (parts.length === 3) {
      const [year, month, day] = parts;
      return `${day}-${month}-${year}`;
    } else {
      logError('Data inválida: ', error)
      return inputDate;
    }
  }

function logMessage(message) {
    const logTable = document.getElementById('log-table');
    const newRow = logTable.insertRow(-1);
    const newCell = newRow.insertCell(0);
    newCell.textContent = message;
}

function logError(message, error) {
    console.error(message, error);
    logMessage(`${message} ${error.message}`);
}