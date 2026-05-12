(function () {
    let template = document.createElement("template");
    template.innerHTML = `
        <style>
            #container { font-family: Arial, sans-serif; padding: 10px; display: flex; flex-direction: column; align-items: center; }
            button { background-color: #0070d2; color: white; border: none; padding: 10px 20px; cursor: pointer; border-radius: 4px; font-size: 14px; }
            button:disabled { background-color: #ccc; cursor: not-allowed; }
            #status { margin-top: 8px; font-size: 12px; color: #666; }
        </style>
        <div id="container">
            <button id="export-btn" disabled>Загрузка системы...</button>
            <div id="status">Инициализация Python (Pandas)...</div>
        </div>
    `;

    class SacExcelWidget extends HTMLElement {
        constructor() {
            super();
            this._shadowRoot = this.attachShadow({ mode: "open" });
            this._shadowRoot.appendChild(template.content.cloneNode(true));
            this._exportBtn = this._shadowRoot.getElementById("export-btn");
            this._status = this._shadowRoot.getElementById("status");
            this._pyodide = null;
        }

        async onCustomWidgetAfterUpdate(changedProperties) {
            if (!this._pyodide && !window.pyodideLoading) {
                window.pyodideLoading = true;
                await this.initPython();
            }
        }

        async initPython() {
            try {
                // 1. Загружаем скрипты из вашего GitHub
                await this.loadScript("https://tokumtaevaz.github.io/my-sac-widget/pyodide.js");
                await this.loadScript("https://tokumtaevaz.github.io/my-sac-widget/xlsx.mini.min.js");

                // 2. Инициализируем Pyodide с указанием пути к файлам WASM
                this._pyodide = await loadPyodide({
                    indexURL: "https://github.io"
                });

                // 3. Загружаем пакет Pandas
                await this._pyodide.loadPackage("pandas");

                this._status.innerText = "Система готова";
                this._exportBtn.innerText = "Скачать график ППР в Excel";
                this._exportBtn.disabled = false;
                this._exportBtn.onclick = () => this.runExport();
            } catch (e) {
                this._status.innerText = "Ошибка загрузки: " + e.message;
            }
        }

        loadScript(src) {
            return new Promise((resolve) => {
                const script = document.createElement("script");
                script.src = src;
                script.onload = resolve;
                document.head.appendChild(script);
            });
        }

        async runExport() {
            this._exportBtn.disabled = true;
            this._status.innerText = "Трансформация данных в Python...";

            try {
                const rawData = this.dataBindings.getDataBinding("p_data").data.map(row => {
                    return {
                        TechPlace: row.dimensions_0.description,
                        Equipment: row.dimensions_1.description,
                        Month: row.dimensions_2.description,
                        WorkType: row.dimensions_3.description
                    };
                });

                this._pyodide.globals.set("raw_json", JSON.stringify(rawData));

                const pythonCode = `
import pandas as pd
import json

data = json.loads(raw_json)
df = pd.DataFrame(data)
month_order = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
df['Month'] = pd.Categorical(df['Month'], categories=month_order, ordered=True)
pivot = df.pivot_table(index=['TechPlace', 'Equipment'], columns='Month', values='WorkType', aggfunc=lambda x: ', '.join(x.unique())).fillna('')
result = pivot.reset_index().to_json(orient='records')
result
                `;

                const pyResult = await this._pyodide.runPythonAsync(pythonCode);
                const finalData = JSON.parse(pyResult);

                const ws = XLSX.utils.json_to_sheet(finalData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "ППР");
                XLSX.writeFile(wb, "График_ППР_SAC.xlsx");

                this._status.innerText = "Готово!";
            } catch (e) {
                this._status.innerText = "Ошибка: " + e.message;
            } finally {
                this._exportBtn.disabled = false;
            }
        }
    }

    customElements.define("com-sap-sac-excel-python", SacExcelWidget);
})();
