var selector;

var SKSCase = {
    "onLoad": function (reinit, alertText, chooseTopicText, chooseModelText) {

        if (reinit) {
            SKSCase.getSystems(alertText, chooseTopicText);
            SKSCase.getModels(alertText, chooseModelText);
        }

        $("#keynumber_required").show();
        $("#model_required").hide();
        $("#serial_required").hide();
        $("#address_required").hide();
        
        $('.system_select').change(function () {
            var val = $(this).val();
            if (val == 0) {
                $('.version_select').empty();
            }
            else {
                SKSCase.getVersions(val, alertText);
                
                if ($('#system option:selected').attr("code") == "ITT") {
                    $("#model_required").show();
                    $("#serial_required").show();
                    $("#address_required").show();
                } else {
                    $("#model_required").hide();
                    $("#serial_required").hide();
                    $("#address_required").hide();
                }
            }
        });
        
        $('#model_select').change(function () {

            if ($('#system option:selected').attr("code") == "ITT") {
                if ($('#model_select option:selected').text() == "_none") {
                    $("#serial_required").hide();
                } else {
                    $("#serial_required").show();
                }
            }
        });
    },
    
    "initMultiSelect": function (maxcount, maxsize, fileSizeAlertMsg)
    {
        selector = new MultiSelector(maxcount, maxsize, fileSizeAlertMsg);
        var input = document.getElementById('file1');
        if (input)
            selector.addElement(input);
    },

    "onCaseDetailLoad" : function(alertText)
    {
        $('tr[data-modal="modal"]').bind("click", function () {
            var id = this.getAttribute("actionid");
            var type = this.getAttribute("actiontype");
            var en = this.getAttribute("actionen");


            $.ajax({
                url: encodeURI("/Ajax/ActionHandler.ashx?actionId=" + id + "&actionEntName=" + en + "&actionType=" + type),
                cache: false,
                async: false,
                dataType: 'text',
                success: function (response) {
                    var result = eval("(" + response + ")");

                    if (result) {
                        if (result.Subject)
                            $("#title").val(result.Subject);
                        if (result.Status)
                            $("#state").val(result.Status);
                        if (result.Owner)
                            $("#owner").val(result.Owner);
                        if (result.Description)
                            $("#description").html(result.Description);
                        if (result.ActionType)
                            $("#actionType").val(result.ActionType);
                        if (result.EmailFrom)
                            $("#from").val(result.EmailFrom);
                        if (result.EmailTo)
                            $("#to").val(result.EmailTo);

                    }
                },
                error: function (ex) {
                    alert(alertText);
                }
            });
        });
    },

    //Функция вынесена на страницу Case.aspx
    "onSave" : function()
    {
        var reqField = ["incidentTitle", "description", "system"];

        if ($('#system option:selected').attr("code") == "ITT") {
            reqField.push("address");
            reqField.push("modelid");
            
            if ($('#model_select option:selected').text() != "_none") {
                reqField.push("serialnumber");
            }
            
        }

        var versions = $('#version');
        if ($('#version option:selected').text() == "" && versions.find("option").length > 1) {
            reqField.push("version");
        }
        
        for (var i = 0; i < reqField.length; i++)
        {
            var field = $('#' + reqField[i]);

            if (field.children().val() == "")
            {
                alert("Поле обязательно для заполения: '" + field.attr("title") + "'");
                return false;
            }
        }

        //Удаление не вложенного элемента input. При его наличии элемент загружается на сервер, чего не должно быть
        var btn = document.getElementById('attachbtn');
        if (btn.input) {
            var id = btn.input.id;

            var rInput = document.getElementById(id);
            if (rInput) {
                var parent = rInput.parentNode;
                parent.removeChild(rInput);
            }
        }

        ShowSpinner("mask2");
        return true;
    },

    "getVersions": function (systemid, alertText) {
        $.ajax({
            url: "/Ajax/CaseFormHandler.ashx?type=version&systemKey=" + systemid,
            cache: false,
            async: false,
            success: function (html) {
                $(".version_select").empty().append(html);
            },
            error: function (ex) {
                alert(alertText);
            }
        });
    },

    "getSystems": function (alertText, chooseTopicText) {

        $.ajax({
            url: "/Ajax/CaseFormHandler.ashx?type=system",
            cache: false,
            async: false,
            success: function (html) {
                $(".system_select").empty().append("<option value=''>" + chooseTopicText + "</option>").append(html);
            },
            error: function (ex) {
                alert(alertText);
            }
        });
    },

    "getModels": function (alertText, chooseModelText) {

        $.ajax({
            url: "/Ajax/CaseFormHandler.ashx?type=model",
            cache: false,
            async: false,
            success: function (html) {
                $(".model_select").empty().append("<option value=''>" + chooseModelText + "</option>").append(html);
            },
            error: function (ex) {
                alert(alertText);
            }
        });
    },

    /* Действия вложения файла.*/
    "attachFile": function () {
        if (selector) {

            if (document.getElementById("fake-file__name").innerHTML == "") {
                return;
            }

            var btn = document.getElementById('attachbtn');
            if (btn.input) {

                if (selector.validate(btn.input)) {

                    // создаем новый input file
                    var new_element = document.createElement('input');
                    new_element.type = 'file';
                    new_element.className = 'fake-file__input';

                    // добавляем перед старым
                    btn.input.parentNode.insertBefore(new_element, btn.input);

                    // добавляем в общий список выбранный файл
                    this.addAttachRow(btn.input);
                    // скрываем старый инпут
                    btn.input.style.position = 'absolute';
                    btn.input.style.left = '-1000px';

                    // добавляем новый инпут
                    selector.addElement(new_element);

                    $('#fake-file__name').empty();

                    //var control = $("#file1");
                    //control.replaceWith(control.clone(true));
                    $("#clearbtn").css('visibility', 'hidden');

                }
            }
        }
        //Не участвует в локализации
        //else
        //    alert("Невозможно произвести процедуру вложения файла.");
    },
    
    /*удаляет и создает на его месте другой input file. Очищает отображение*/
    "cloneInputFile" : function(rInput)
    {
        if (rInput) {
            var id = rInput.id;

            var parent = rInput.parentNode;
            parent.removeChild(rInput);

            var nInput = document.createElement('input');
            nInput.id = id;
            nInput.name = id;
            nInput.type = 'file';
            nInput.className = 'fake-file__input';

            parent.appendChild(nInput);
            nInput.onchange = function () { sub(this); };

            document.getElementById("fake-file__name").innerHTML = "";
            $("#clearbtn").css('visibility', 'hidden');

            return nInput;
        }
    },

    /* Действия вложения файла.*/
    "dropFile": function () {
        if (selector) {
                var btn = document.getElementById('attachbtn');
                if (btn.input) {

                    var nInput = SKSCase.cloneInputFile(btn.input);
                    btn.input = nInput;
                }
        }
        //Не участвует в локализации
        //else
        //    alert("Невозможно произвести процедуру удаления файла.");
    },

    /*Add a new row to the list of files*/
    "addAttachRow": function (element) {

        var fullPath = element.value;
        if (fullPath) {
            var startIndex = (fullPath.indexOf('\\') >= 0 ? fullPath.lastIndexOf('\\') : fullPath.lastIndexOf('/'));

            var filename = fullPath.substring(startIndex);

            if (filename.indexOf('\\') === 0 || filename.indexOf('/') === 0) {
                filename = filename.substring(1);
            }

            var fileList = $('#files_list');

            var li = $('<li><span class="name">' + filename + ' </span><span class="del"></span></li>');
            fileList.append(li);

            li[0].element = element;

            $('.del').click(function () {
                var li = $(this).closest('li')[0];
                var currentInput = element;
                selector.removeSize(currentInput);

                selector.count--;
                selector.current_element.disabled = false;

                $(li.element).remove();
                $(li).remove();

                return false;
            });
        }
    }
};

