'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#addSection').click(addSection);
            $('#saveSections').click(saveSections);
        });
    });


    function addSection() {
        let uid = String(genUID())
        let new_row = document.createElement('tr')
        new_row.id = uid + "_row"
        let td_name = document.createElement('td')
        td_name.setAttribute("contenteditable", "true")
        td_name.id = uid + "_name"
        let td_base = document.createElement('td')
        td_base.setAttribute("contenteditable", "true")
        td_base.id = uid + "_base"
        let td_height = document.createElement('td')
        td_height.setAttribute("contenteditable", "true")
        let td_erase = document.createElement('td')
        td_height.id = uid + "_height"
        new_row.appendChild(td_base)
        new_row.appendChild(td_height)
        new_row.appendChild(td_name)
        new_row.appendChild(td_erase)
        let delButton = document.createElement('button')
        delButton.innerText = "X"
        delButton.id = uid
        delButton.addEventListener("click", function () {
            document.getElementById(this.id + "_row").remove()
        });
        td_base.addEventListener("keyup", function () {
            updateSectionName(uid)
        })
        td_height.addEventListener("keyup", function () {
            updateSectionName(uid)
        })
        td_erase.appendChild(delButton)
        document.getElementById("sections_table").appendChild(new_row)
        td_base.focus()
    }

    function saveSections() {
        // Que no se pueda enviar una seccion con algun campo vacio ni secciones con el mismo nombre
        //eraseAllSections()
        Excel.run(function (context) {
            //var range = context.workbook.getSelectedRange();
            //range.format.fill.color = 'green';
            //range.values = [["Hola"]]
            //return context.sync();
            var table = document.getElementById('sections_table')
            var sheet = context.workbook.worksheets.getItem("Entrada");
            var headings = [
                ["Secciones", "b", "h"],
            ];
            var content = []
            var base = 0
            var height = 0
            var name = ""
            var id = ""
            var range = sheet.getRange("A26:C26")
            range.values = headings;
            //range.format.autofitColumns();
            for (var i = 1, row; row = table.rows[i]; i++) {
                id = String(String(row.id).match(/[0-9]*/i)[0]) // Match the id
                base = document.getElementById(id + "_base").innerHTML
                height = document.getElementById(id + "_height").innerHTML
                name = document.getElementById(id + "_name").innerHTML
                content.push([name, base, height])
            }
            var content_range = sheet.getRangeByIndexes(26, 0, table.rows.length - 1, 3)
            content_range.values = content;
            return context.sync();
        }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function updateSectionName(uid) {
        document.getElementById(uid + "_name").innerHTML = document.getElementById(uid + "_base").innerHTML + "x" + document.getElementById(uid + "_height").innerHTML;
    }

    function eraseAllSections() {
        Excel.run(function (context) {
            var sheet = context.workbook.worksheets.getItem("Entrada");
            var row = 26
            var check_range = sheet.getRangeByIndexes(row, 0, 1, 1)
            while (check_range.values != "") {
                full_range = sheet.getRangeByIndexes(row, 0, 1, 3)
                full_range.values = [["", "", ""]]
                check_range = sheet.getRangeByIndexes(row, 0, 1, 1)
                row += 1
                context.sync();
            }
            return context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function genUID() {
        return Date.now()
    }

})();
