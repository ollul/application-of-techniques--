
function Check(input) 
{
    input.value = input.value.replace(/\D/g, "");
}

function changeButtonColor(obj)  //Цвета кнопок
{
	obj.className = (obj.className == 'bgltorange-fntblue') ? 'bgltgrey-fntblue' : 'bgltorange-fntblue';
}

$(function() {
	
    for (var i = 0; i < dor_array.length; i++) {
        $("<option value='" + dor_array[i][0] + "'>" + dor_array[i][0] + " - " + dor_array[i][1] + "</option>").appendTo("#dor-select");
    }

    $("<option value='0'>Все</option>").appendTo("#depo-select");

    for (var i = 0; i < depo_array.length; i++) {
        if (depo_array[i][2] == 1) {
            $("<option value='" + depo_array[i][0] + "'>" + depo_array[i][0] + " - " + depo_array[i][1] + "</option>").appendTo("#depo-select");
        }
    }
	
	$("<option value='0'>Все</option>").appendTo("#skp-select");
    
    for (var i = 0; i < skp_array.length; i++) {
    	$("<option value='" + skp_array[i][2] + "'>" + skp_array[i][0] + " - " + skp_array[i][1] + "</option>").appendTo("#skp-select");
    } 
	
    $("#dor-select").change(function() {
		
        $("#depo-select").empty();
		
        $("<option value='0'>Все</option>").appendTo("#depo-select");
        for (var i = 0; i < depo_array.length; i++) {
            if (depo_array[i][2] == $("#dor-select").val()) {
                $("<option value='" + depo_array[i][0] + "'>" + depo_array[i][0] + " - " + depo_array[i][1] + "</option>").appendTo("#depo-select");
            }
        }
     });

	$('#otch-select').change(function() {
		
      if ($(this).val() == '1') {
        $('#date_from').css('display', 'block');
        $('#date_to').css('display', 'block');
        $('#year-select').css('display', 'none');
        $('#month-select').css('display', 'none');        
      } else {
        $('#year-select').css('display', 'block');
        $('#month-select').css('display', 'block');
        $('#date_from').css('display', 'none');
        $('#date_to').css('display', 'none');
      }
    });
	
	$('#tabn-input, #nagon-input, #dtrain-input, #dtrain1-input').keydown(function(event) {
        if (event.keyCode == 46 || event.keyCode == 8 || event.keyCode == 9 || event.keyCode == 27 ||
            (event.keyCode == 65 && event.ctrlKey === true) ||
            (event.keyCode >= 35 && event.keyCode <= 39)) {
            return;
        } else if ((event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105)) {
            event.preventDefault();
        }
    });
	
	var month = {
		1 : "Январь",
		2 : "Февраль",
		3 : "Март",
		4 : "Апрель",
		5 : "Май",
		6 : "Июнь",
		7 : "Июль",
		8 : "Август",
		9 : "Сентябрь",
		10 : "Октябрь",
		11 : "Ноябрь",
		12 : "Декабрь"
    }

  for(var i = new Date().getMonth() + 1; i >= 1 ; i--) {
    $('#month-select').append("<option value='"+ i +"'>" + month[i] + "</option>");
  }
  
  $('#year-select').change(function() {    
    $('#month-select').empty();
    if (new Date().getFullYear() == $('#year-select').val()) {
      for(var i = new Date().getMonth() + 1; i >= 1 ; i--) {
        $('#month-select').append("<option value='"+ i +"'>" + month[i] + "</option>");
      }
    } else {
      for(var i = 12; i >= 1 ; i--) {
        $('#month-select').append("<option value='"+ i +"'>" + month[i] + "</option>"); 
      }
    }    
  });
	
});

function addOption(idSelect, val, txt)                                  // функция, формирующая набор options
{ 
	var OptionNew = document.createElement('option');
	OptionNew.value = val;
	OptionNew.title = txt;
	OptionNew.text  = txt;
	try 
	{
		idSelect.add(OptionNew,null);
	}
	catch(ex) 
	{
		idSelect.add(OptionNew);
	}
}

function dtHTML2dtDB2(sdate)                                            // переводит дату в нужный формат
{
	/* dd.mm.yyyy -> yyyy-mm-dd */
	var sdt = sdate.substring(6,10) + '-' + sdate.substring(3,5) + '-' + sdate.substring(0,2); 
	return sdt; 
}

function showReport()             // при нажатии "показать справку"
{
	if( imgRptView.className == 'img' )
	{
		showParams();
		imgExcel.className = 'img';
		imgPrint.className = 'img';	
		DoExel.disabled=false;
		DoPrint.disabled=false;
		divContentH.style.visibility = 'hidden';
		divContentR.style.visibility = 'hidden';
		divWaitShow(1);
		formRepHref(); 
	}
}

 function correct_data(year, month, day) {
   if (month.length == 1) month = '0' + month;
   if (day.length == 1) day = '0' + day;
   return year + '-' + month + '-' + day;
}

function formRepHref()           // формирование ссылки с параметрами
{
	var date1 = new Date($('#date_from').val().substr(6, 4), parseFloat($('#date_from').val().substr(3, 2), 10) - 1, $('#date_from').val().substr(0, 2));
    var date2 = new Date($('#date_to').val().substr(6, 4), parseFloat($('#date_to').val().substr(3, 2), 10) - 1, $('#date_to').val().substr(0, 2)); 
    var one_day = 1000 * 60 * 60 * 24;
    var rez = Math.floor((date2 - date1) / one_day) + 1; // разница для 3 мес
	
	if ($('#dor-select').val() != 0 && rez > 93 && $('#otch-select').val() == 3) {
		alert('По всей сети период ограничен 3 месяцами');
		return;
    }
	
	if ($('#dor-select').val() != -1 && $('#depo-select').val() == 0 && rez > 93) {
		alert('По дороге период ограничен 3 месяцами');
		return;
    }
	
    if ($('#dor-select').val() != -1 && $('#depo-select').val() != 0 && rez > 93) {
		alert('По депо период ограничен 3 месяцами');
		return;
    }
	
	if ($('#nagon-input').val() == 0) {
		alert('Введите нагон >= 1 минуты');
		return;
    }
   
    /* repHref */
	
	var repUrl = document.location.href.substr( 0, document.location.href.indexOf('?') ) + '?_program=/OCRVFoundation/AOM/Daily/SP/nagon_dor&rpt=1';

	var repHref = '';
	repHref += repUrl;
	
	var time = (new Date()).getTime(); 
	repHref += '&' + 'time=' + time;
	repHref += '&' + 'road=' + $('#dor-select').val();
	repHref += '&' + 'depo=' + $('#depo-select').val();
	repHref += '&' + 'tab_num=' + ($('#tabn-input').val() != '' ? $('#tabn-input').val() : 0); 
	repHref += '&' + 'skp=' + $('#skp-select').val();
    repHref += '&' + 'doljn=' + $('#doljn-select').val(); 
	repHref += '&' + 'train_num=' + ($('#train-input').val() != '' ? $('#train-input').val() : 0); 
	repHref += '&' + 'train_num_beg=' + ($('#dtrain-input').val() != '' ? $('#dtrain-input').val() : 0);
	repHref += '&' + 'train_num_end=' + ($('#dtrain1-input').val() != '' ? $('#dtrain1-input').val() : 0);
	repHref += '&' + 'nagon=' +  ($('#nagon-input').val() != '' ? $('#nagon-input').val() : 0);
	repHref += '&' + 'vd=' + $('#vd-select').val();
	repHref += '&' + 'itogi=' + ($('#itogi').is(':checked') ? 1 : 0); // итоги
	
	var lastDay = (new Date($('#year-select').val(), $('#month-select').val(), 0)).getDate();
	
	if ($('#otch-select').val() == '1') {
		repHref += '&' + 'otch=1';  /*оперативная отчетность*/
		repHref += '&' + 'dt1=' + dtHTML2dtDB2($('#date_from').val());
		repHref += '&' + 'dt2=' + dtHTML2dtDB2($('#date_to').val());
    } else {
		repHref += '&' + 'otch=2';  /*статистическая отчетность*/
		repHref += '&' + 'dt1=' + correct_data($('#year-select').val(), $('#month-select').val(), '1');
		repHref += '&' + 'dt2=' + correct_data($('#year-select').val(), $('#month-select').val(), lastDay);
    }
	
	repHref += '&' + 'year=' + $('#year-select').val();
    repHref += '&' + 'month=' + $('#month-select').val();
  	
	$('.skp').click(function() {
		window.open(document.location.href.substr(0, document.location.href.indexOf('?')) + $(this).attr('value'));
    });
    $('.skp').hover(function() { $(this).css({'color': '#027ffc', 'cursor': 'pointer'}); }, function() { $(this).css({'color': 'black', 'cursor': 'default'}); });
	
	$('.tab_nagon').click(function() {
		window.open(document.location.href.substr(0, document.location.href.indexOf('?')) + $(this).attr('value'));
    });
    $('.tab_nagon').hover(function() { $(this).css({'color': '#027ffc', 'cursor': 'pointer'}); }, function() { $(this).css({'color': 'black', 'cursor': 'default'}); });
	
	
	var s='';
	
    if ($('#otch-select').val() == '1') {
      s+='<label class=fltkey>Отчетность:</label><label class=fltval>Оперативная</label> ';
      s+='<label class=fltkey>Период:</label><label class=fltval>'
        +rptDt1.value+' - '+rptDt2.value+'</label> ';
    } else {
      s+='<label class=fltkey>Отчетность:</label><label class=fltval>Статистическая</label> ';
      s+='<label class=fltkey>Месяц:</label><label class=fltval>'
        +$('#month-select option:selected').text()+' '+$('#year-select option:selected').text()+'г.  '+'</label> ';
    }
    s+='<label class=fltkey>Дорога:</label><label class=fltval>'
		+$('#dor-select option:selected').text()+'</label> ';
	s+='<label class=fltkey>Депо:</label><label class=fltval>'
		+$('#depo-select option:selected').text()+'</label> ';
	s+='<label class=fltkey>Вид движения:</label><label class=fltval>'
		+$('#vd-select option:selected').text()+'</label><br>';
	s+='<label class=fltkey>Таб номер раб.:</label><label class=fltval>'
		+($('#tabn-input').val() != 0 ? $('#tabn-input').val() : 'Все')+'</label> ';
    s+='<label class=fltkey>Сетевой код перев.:</label><label class=fltval>'
		+$('#skp-select option:selected').text()+'</label> ';
    s+='<label class=fltkey>Должность:</label><label class=fltval>'
		+$('#doljn-select option:selected').text()+'</label> '; 
	s+='<label class=fltkey>Нагон &ge;:</label><label class=fltval>'
		+($('#nagon-input').val() != '' ? $('#nagon-input').val() : 0)+'</label> ';
   /* s+='<label class=fltkey>Номер поезда:</label><label class=fltval>'
		+($('#train-input').val() != '' ? $('#train-input').val() : 0)+'</label> '; */
    s+='<label class=fltkey>Диапазон номеров поездов:</label><label class=fltval>'
	    +($('#dtrain-input').val() != 0 ? $('#dtrain-input').val() : 'Все') + ' - '  
		+($('#dtrain1-input').val() != 0 ? $('#dtrain1-input').val() : 'Все')+'</label> ';
    tdParams.innerHTML=s;      
	   /* alert(repHref); */
	/* данные отчета */
    startDownload(endDownload, repHref);    
}

function startDownload(callBack,url)              // начало загрузки
{
	if (window.XMLHttpRequest)
	{
        try
		{
            req = new XMLHttpRequest();
        }
		catch (e){}
    } 
	else if (window.ActiveXObject)
	{
        try
		{
            req = new ActiveXObject('Msxml2.XMLHTTP');
        }
		catch (e)
		{
            try 
			{
                req = new ActiveXObject('Microsoft.XMLHTTP');
            }
			catch (e){}
        }
	}
	if (req)
	{
		req.open('GET',url,true);
		req.onreadystatechange=function()  { if (req.readyState==4) callBack(req,url);};
		req.send(null);
	}
	else alert('Браузер не поддерживает технологию AJAX');
}

function endDownload(content,url)               //конец
{
	divWaitHide();
	if (content.status != 200)
	{
		/* unsuccessful */ 
		alert('Ошибка: content.status=' + content.status + '!');
		return;
	}
	if (content.responseText.indexOf('SASLogbutton')!= -1) 
	{ 
		/* ошибки sas */
		window.location.href = url;
		return; 
	} 
	if (content.responseText.indexOf('SASLogon') != -1) 
	{ 
		/* коннекция отвалилась по таймайту, обновляем страницу */ 
		window.location.reload(); 
		return; 
	}
	i=content.responseText.indexOf('@');
	divContentH.innerHTML=content.responseText.substring(0,i);
	divContentR.innerHTML=content.responseText.substring(i+1);
	document.body.focus();
	window.status = "Готово";
	afterDownload();                                          // function afterDownload() - своя функция для каждого отчета: погасить кнопки при отсутствии данных и пр
}

function afterDownload()         
{
	if (document.getElementById('rptTblM') != null)
	{
		if (rptTblM.rows.length < 1)
		{
	        addRow2Table('rptTblM', 'Нет данных', 'tdNoData', 14);
			hideExPr();
		}
		divContentH.style.visibility = 'visible';
		divContentR.style.visibility = 'visible';
	}
}

function addRow2Table(tblName,tdHTML,tdClass,tdColSpan)    //добавление строки
{
	var tbody = document.getElementById(tblName).getElementsByTagName('TBODY')[0];
	row = document.createElement('TR');
	tbody.appendChild(row);
	var td = document.createElement('TD');
	td.colSpan=tdColSpan;
	row.appendChild(td); 
	td.innerHTML = tdHTML; 
	td.className = tdClass;
}

function hideExPr()   // скрытие кнопок печати и экспорта в эксель
{
	imgExcel.className='imgD';
	imgPrint.className='imgD';	
	DoExel.disabled=true;
	DoPrint.disabled=true;
}

function OpenExcel()   // открытие Exel
{ 
	divWaitShow(0); 
	setTimeout(doSave, 500); 
} 

function doSave()   // попытка создать Exel
{ 
	try { ExlPrn(0); }                                                  // function ExlPrn(prn) - своя функция для каждого отчета, prn=1 - печать 
	catch(e) 
	{  
		alert( 'Ошибка [' + ( e.number & 0xFFFF ) + '] ' + e.description ); 
		divWaitHide();
	} 
} 

function doPrint() 
{  
	
	try { ExlPrn(1); }	
	catch(e) 
	{  
		alert( 'Ошибка [' + ( e.number & 0xFFFF ) + '] ' + e.description ); 
		divWaitHide();
	} 
}

function ExlPrn(prn)
{
	if (imgExcel.className == 'img')
	{
		var app = new ActiveXObject('Excel.Application');
		var bk  = app.Workbooks.Add();
		app.SheetsInNewWorkBook = 1;
		var sh = bk.Worksheets(1);
		sh.Select;
		sh.PageSetup.Zoom = false;
		sh.PageSetup.FitToPagesWide = 1;
		sh.PageSetup.FitToPagesTall = 999;
		sh.PageSetup.Orientation = 2;		
		sh.PageSetup.PrintTitleRows = "$3:$4";
		sh.PageSetup.CenterHorizontally = true;		

		sh.PageSetup.LeftMargin = app.CentimetersToPoints(1);           // отступы (в см)
		sh.PageSetup.RightMargin = app.CentimetersToPoints(1);
		sh.PageSetup.TopMargin = app.CentimetersToPoints(1.5);
		sh.PageSetup.BottomMargin = app.CentimetersToPoints(1.5);

		sh.PageSetup.RightHeader = 'Автоматизированное формирование дорожной и аналитической отчетности ЦОММ на сетевом уровне';
		sh.PageSetup.RightFooter = '&D &N';                             // заполнение нижнего колонтитула
		sh.PageSetup.LeftFooter = 'Справка о нагоне опозданий пассажирских и пригородных поездов   &P/&N';
		
		sh.Cells(1, 1) = 'Справка о нагоне опозданий пассажирских и пригородных поездов ';
		sh.Range('A1:C1').MergeCells = true;
		sh.Range('A1:C1').Select;
		sh.Rows('1:1').RowHeight = 35;
		app.Selection.HorizontalAlignment = -4108;
		sh.Cells(1, 1).Characters.Font.Name = 'Arial Black';
		sh.Cells(1, 1).Characters.Font.Size = 12;

		sh.Cells(2, 1) = tdParams.innerText;
		sh.Range('A2:C2').MergeCells = true;
		sh.Range('A2:C2').Select;
		sh.Rows('2:2').RowHeight = 35;
		app.Selection.HorizontalAlignment = -4108;
		sh.Cells(2, 1).Characters.Font.Name = 'Arial';
		sh.Cells(2, 1).Characters.Font.Size = 10;
		

		window.clipboardData.setData('Text','<table>' + rptTblH.innerHTML + '</table>');
		sh.Paste(sh.Cells(3, 1));
		sh.Rows('3:3').RowHeight = 60;
		sh.Range('A3:C3').Select;
		app.Selection.Interior.Color = 13535777;
		try{
		app.Selection.Font.ThemeColor = 1;
		}
		catch(e)
		{
				app.Application.Visible = true;
		}
		app.Selection.Font.TintAndShade = 0;
		
		sh.Columns("A:A").ColumnWidth = 55;
		sh.Columns("B:B").ColumnWidth = 40;
		sh.Columns("C:C").ColumnWidth = 40;	
	}
	
	var method=0;

	if (document.getElementById('rptTblM') != null)
	{
		if (rptTblM.rows.length >= 1)
		{	
			try{
				window.clipboardData.setData('Text','<table>' + rptTblM.innerHTML + '</table>');
				sh.Paste(sh.Cells(4, 1));
			}
			catch(e)
			{
				method=1;
			}
			if(method==1)
			{				
				sh.Range('A3', sh.Cells(sh.Cells.SpecialCells(11).Row, sh.Cells.SpecialCells(11).Column)).Select;
				app.Selection.VerticalAlignment = -4108;
				app.Selection.Font.Name = 'Arial';
				app.Selection.Font.Size = 10;
				try{
				app.Selection.Borders(7).LineStyle = 1;
				app.Selection.Borders(8).LineStyle = 1;
				app.Selection.Borders(9).LineStyle = 1;
				app.Selection.Borders(10).LineStyle = 1;
				app.Selection.Borders(11).LineStyle = 1;
				app.Selection.Borders(12).LineStyle = 1;
				}
				catch(e)
				{
					app.Application.Visible = true;
				}
				sh.Columns("B:B").Select;
				app.Selection.HorizontalAlignment = -4108;
				
				sh.Rows("4:4").EntireRow.AutoFit;
				sh.Rows("3:3").EntireRow.AutoFit;

				sh.Range("A1:C1").Select;	
			}
			else
			{
				sh.Range('A3', sh.Cells(sh.Cells.SpecialCells(11).Row, sh.Cells.SpecialCells(11).Column)).Select;
				app.Selection.VerticalAlignment = -4108;
				app.Selection.Font.Name = 'Arial';
				app.Selection.Font.Size = 10;
				try{
				app.Selection.Borders(7).LineStyle = 1;
				app.Selection.Borders(8).LineStyle = 1;
				app.Selection.Borders(9).LineStyle = 1;
				app.Selection.Borders(10).LineStyle = 1; 
				app.Selection.Borders(11).LineStyle = 1;
				app.Selection.Borders(12).LineStyle = 1;
				}
				catch(e)
				{
					app.Application.Visible = true;
				}
				sh.Columns("B:B").Select;
				app.Selection.HorizontalAlignment = -4108;
				
				sh.Rows("4:4").EntireRow.AutoFit;
				sh.Rows("3:3").EntireRow.AutoFit;

				sh.Range("A1:C1").Select;
			}
		}
		divWaitHide();

		if (prn == 1)
		{
			if(method==1)
			{
				setTimeout( function () {sh.PrintOut();
				sh.Application.DisplayAlerts = false;
				sh.Application.Quit(); },1500);
			}
			else
			{
				sh.PrintOut();
				sh.Application.DisplayAlerts = false;
				sh.Application.Quit();
			}
		}
		else
		{
			if(method==1)
			{
				setTimeout( function () {
				try { app.Application.ActiveWindow.View = 3; }
				catch(e){}
			
				app.Application.Visible = true; },1500);
			}
			else
			{
				try { app.Application.ActiveWindow.View = 3; }
				catch(e){}
			
				app.Application.Visible = true;
			}
		}
	}
}

function PrintReport() 
{ 
	divWaitShow(0); 
	setTimeout(doPrint, 500); 
} 

/*function showNagon(dt1, dt2, otch, road, depo, tab_num, skp, doljn, train_num, train_num_beg, train_num_end, nagon, vd)
{
	url = URLproval + dt1 + dt2 + otch+ road + depo + tab_num + skp + doljn + train_num + train_num_end + nagon + vd;
	window.open(url);
}
*/

function showNagon(dt1, dt2, otch, road, depo, tab_num, skp, doljn, train_num, train_num_beg, train_num_end, nagon, vd, itogi, year, month)
{
	url = URLproval + dt1 + dt2 + otch + road + depo + tab_num + skp + doljn + train_num + train_num_beg + train_num_end + nagon + vd + itogi + year + month;
	window.open(url);
}

