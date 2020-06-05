
$(document).ready(
    function() {
        $("#btnRecover").click(Remind);
        $("#btnClose").click(Close);
        $("#txtPhone").mask("+7 (999) 999-99-99", { noshift: true });
    }
);

//Функция перенесена в статический локализованный контент RemindPasswordContentRu.html\RemindPasswordContentEn.html
function Remind() {
    
    $(".recovery__no").hide();
    $(".recovery__ok").hide();
    $("#spanPhone").removeClass("enter__input-err");
    $("#spanEmail").removeClass("enter__input-err");
    $("#spanCode").removeClass("enter__input-err");

    var phone = $("#txtPhone").val();
    var email = $("#txtEmail").val();
    var code = $("#txtCode").val();

    if ( code == "") {
        $("#spanPhone").addClass("enter__input-err");
        $("#spanEmail").addClass("enter__input-err");
        $(".recovery__no__par").text("Не указан персональный код.");
        $(".recovery__no").show();
        return;
    }
    
    if (phone == "" && email == "") {
        $("#spanPhone").addClass("enter__input-err");
        $("#spanEmail").addClass("enter__input-err");
        $(".recovery__no__par").text("Не указаны реквизиты для отправки регистрационных данных.");
        $(".recovery__no").show();
        return;
    }

    ShowSpinner();
    $.ajax({
        url: encodeURI("/Ajax/Remindservice.ashx?code=" + code + "&email=" + email + "&phone=" + phone),
        cache: false,
        async: true,
        dataType: 'text',
        success: function (response) {
            HideSpinner();
            var result = eval("(" + response + ")");
            if (result) {
                if (result.IsPersonalCodeValid)
                {
                    if (result.IsEmailValid || result.IsPhoneValid)
                    {
                        // v Ваши регистрационные данные отправлены указанными Вами способом доставки
                        $(".recovery__ok").show();
                        
                        $("#txtPhone").val("");
                        $("#txtEmail").val("");
                        $("#txtCode").val("");
                    }
                    else
                    {
                        if (phone != "")
                            $("#spanPhone").addClass("enter__input-err");
 
                        if(email != "")
                            $("#spanEmail").addClass("enter__input-err");
                        
                        $(".recovery__no__par").text("Адрес электронной почты и/или мобильный телефон не совпадают с данными, указанными при регистрации");
                        $(".recovery__no").show();
                        
                        //$("#txtPhone").val("");
                        //$("#txtEmail").val("");
                        //$("#txtCode").val("");
                        //$("#spanPhone").removeClass("enter__input-err");
                        //$("#spanEmail").removeClass("enter__input-err");
                        //$("#spanCode").removeClass("enter__input-err");
                    }
                }
                else
                {
                    //enter__input-err
                    $("#spanCode").addClass("enter__input-err");
                    $(".recovery__no__par").text("Указанный Персональный код не найден");
                    $(".recovery__no").show();
                    
                    //$("#txtPhone").val("");
                    //$("#txtEmail").val("");
                    //$("#txtCode").val("");
                    //$("#spanPhone").removeClass("enter__input-err");
                    //$("#spanEmail").removeClass("enter__input-err");
                    //$("#spanCode").removeClass("enter__input-err");
                    
                }
            }
        },
        error: function (ex) {
            HideSpinner();
            alert('Ошибка обработки запроса.');
        }
    });
    //Отправляем
    
}

function Close() {
    document.location.href = "login.aspx";
}