async function create_spreadsheet_from_clipboard() {
    try {
        const dateInput = document.getElementById('date-input').value;
        let options = {
            mode: 'text',
            pythonPath: window.python.path('../.venv/Scripts/python.exe'),
            args: ['create_spreadsheet_from_clipboard', dateInput]
        }
        const scriptPath = window.python.path('../main.py');

        await window.python.runPython(scriptPath, options);
    } catch (error) {
        console.error('Erro durante a execução do script Python:', error);
    }
}

async function generate_spreadsheets() {
    try {
        const dateInput = document.getElementById('date-input').value;
        let options = {
            mode: 'text',
            pythonPath: window.python.path('../.venv/Scripts/python.exe'),
            args: ['generate_spreadsheets', dateInput]
        }
        const scriptPath = window.python.path('../main.py');

        await window.python.runPython(scriptPath, options);
    } catch (error) {
        console.error('Erro durante a execução do script Python:', error);
    }
}

async function send_email() {
    try {
        let options = {
            mode: 'text',
            pythonOptions: ['-u'],
            pythonPath: window.python.path('../.venv/Scripts/python.exe'),
            args: ['send_email']
        }
        const scriptPath = window.python.path('../main.py');

        await window.python.runPython(scriptPath, options);
    } catch (error) {
        console.error('Erro durante a execução do script Python:', error);
    }
}