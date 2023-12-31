const { contextBridge } = require('electron')
let { PythonShell } = require('python-shell')
const os = require('os')
const path = require('path')

contextBridge.exposeInMainWorld('python', {
  homedir: () => os.homedir(),
  path: (joinStr) => path.join(__dirname, joinStr),
  runPython: (name, options) => {
    PythonShell.run(name, options);
  }
})

window.addEventListener('DOMContentLoaded', () => {
    const replaceText = (selector, text) => {
      const element = document.getElementById(selector)
      if (element) element.innerText = text
    }
  
    for (const dependency of ['chrome', 'node', 'electron']) {
      replaceText(`${dependency}-version`, process.versions[dependency])
    }
  })
