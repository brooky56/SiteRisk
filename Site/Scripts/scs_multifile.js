
function MultiSelector(max, maxsize, alertFileSizeText) {

    this.count = 0;

    this.id = 0;

    this.currentsize = 0;

    if (max) {
        this.max = max;
    } else {
        this.max = -1;
    };
  
    if (maxsize) {
        this.maxsize = maxsize;
    } else {
        this.maxsize = -1;
    };

    if (alertFileSizeText) {

        this.alertMsg = alertFileSizeText;
    }

    this.btn = document.getElementById('attachbtn');

    this.validate = function (input) {
        // получение файла. проверка общего размера
        var file;
        if (navigator.appName.indexOf("Internet Explorer") != -1) {
            try {
                var myFSO = new ActiveXObject("Scripting.FileSystemObject");
                file = myFSO.getFile(input.value);
            } catch(e) {

            } 

        }
        else {
            var file = input.files[0];
        }

        if (file != null) {
            if (this.currentsize + file.size > this.maxsize) {
                alert(alertFileSizeText + " " + this.maxsize / (1024 * 1024) + "Mb");
                return false;
            } else
                this.currentsize += file.size;
        }
        return true;
    },

    /*добавление информации в новый элемент*/
    this.addElement = function (element) {

        // Make sure it's a file input element
        if (element.tagName == 'INPUT' && element.type == 'file') {

            element.name = 'file_' + this.id++;
            element.id = element.name;

            this.btn.input = element;

            element.onchange = function () {
                sub(this);
            };

            // If we've reached maximum number, disable input element
            if (this.max != -1 && this.count >= this.max) {
                element.disabled = true;
            };

            // File element counter
            this.count++;
            // Most recent element
            this.current_element = element;

        } else {
            // This can only be applied to file input elements!
            alert('Error: not a file input element');
        };

    };

    
    /*добавление информации в новый элемент*/
    this.removeSize = function (input) {
        // получение файла. проверка общего размера
        var file;
        if (navigator.appName.indexOf("Internet Explorer") != -1) {
            try {
                var myFSO = new ActiveXObject("Scripting.FileSystemObject");
                file = myFSO.getFile(input.value);
            } catch (e) {

            }

        }
        else {
            file = input.files[0];
        }

        if (file != null) {
            this.currentsize -= file.size;
        }
    };
};