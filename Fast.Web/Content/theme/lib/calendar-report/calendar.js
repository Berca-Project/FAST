
const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

var renderCJan = 0;

//initiate Feb
var renderCFeb = 0;

//initiate Mar
var renderCMar = 0;

//initiate Apr
var renderCApr = 0;

//initiate May
var renderCMay = 0;

//initiate June
var renderCJun = 0;

//initiate July
var renderCJul = 0;

//initiate August
var renderCAug = 0;

//initiate Sept
var renderCSept = 0;

//initiate Oct
var renderCOct = 0;

//initiate Nov
var renderCNov = 0;

//initiate December
var renderCDec = 0;

function printDiv(divName) {

	var printContents = document.getElementById(divName).innerHTML;
	var originalContents = document.body.innerHTML;

	document.body.innerHTML = printContents;
	window.print();

	document.body.innerHTML = originalContents;
}

function collectCalendar() {

	var tableHoliday = $('#holidayTable').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"info": false,
		"filter": false,
		"bLengthChange": false,
		"orderMulti": false,
		"pageLength": 50,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllHolidayWithParam",
			"data": { locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs": [
			{ "targets": [0], "searchable": false, "visible": false, "orderable": false },
			{ "targets": [1], "searchable": true, "orderable": false },
			{ "targets": [2], "searchable": true, "orderable": false },
			{ "targets": [3], "searchable": true, "orderable": false }
		],
		"columns": [
			{ "data": "ID", "name": "ID", "autoWidth": false },
			{
				"data": "Date", "Title": "Bulan",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("MMMM");
					else
						return '';
				}
			},
			{
				"data": "Date", "Title": "Tanggal",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{ "data": "Description", "name": "Description", "Title": "Informasi", "autoWidth": true }
		],
		"fnRowCallback": function (nRow, data, iDisplayIndex, iDisplayIndexFull) {
			$('td', nRow).css('background-color', data['Color']);
			$('#holidayTable_paginate')[0].style.display = "none";
		}
	});

	var workingSummaryTable = $('#workingSummaryTable').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"info": false,
		"filter": false,
		"bLengthChange": false,
		"orderMulti": false,
		"pageLength": 25,
		"destroy": true,
		"bSort": false,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetWorkingSummaryWithParam",
			"data": { locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs": [
			{ "targets": [0], "searchable": false, "orderable": false },
			{ "targets": [1], "searchable": false, "orderable": false },
			{ "targets": [2], "searchable": false, "orderable": false },
			{ "targets": [3], "searchable": false, "orderable": false },
			{ "targets": [4], "searchable": false, "orderable": false },
			{ "targets": [5], "searchable": false, "orderable": false },
			{ "targets": [6], "searchable": false, "orderable": false }
		],
		"columns": [
			{ "data": "ColumnName", "name": "ColumnName", "Title": "2020", "autoWidth": false },
			{ "data": "Days", "name": "Days", "autoWidth": false },
			{ "data": "Holiday", "name": "Holiday", "autoWidth": false },
			{ "data": "Leaves", "name": "Leaves", "autoWidth": false },
			{ "data": "ProdOff", "name": "ProdOff", "autoWidth": false },
			{ "data": "ShiftOff", "name": "ShiftOff", "autoWidth": false },
			{ "data": "WorkDays", "name": "WorkDays", "autoWidth": false }
		],
		"fnDrawCallback": function () {
			$('#workingSummaryTable_paginate')[0].style.display = "none";
		}
	});

	var workingGroupSummaryTable = $('#workingGroupSummaryTable').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"info": false,
		"filter": false,
		"bLengthChange": false,
		"orderMulti": false,
		"pageLength": 25,
		"destroy": true,
		"bSort": false,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetWorkingGroupSummaryWithParam",
			"data": { locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs": [
			{ "targets": [0], "searchable": false, "orderable": false },
			{ "targets": [1], "searchable": false, "orderable": false },
			{ "targets": [2], "searchable": false, "orderable": false },
			{ "targets": [3], "searchable": false, "orderable": false },
			{ "targets": [4], "searchable": false, "orderable": false }
		],
		"columns": [
			{ "data": "ColumnName", "name": "ColumnName", "Title": "2020", "autoWidth": false },
			{ "data": "A", "name": "A", "Title": "A", "autoWidth": true },
			{ "data": "B", "name": "B", "Title": "B", "autoWidth": true },
			{ "data": "C", "name": "C", "Title": "C", "autoWidth": true },
			{ "data": "D", "name": "D", "Title": "D", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#workingGroupSummaryTable_paginate')[0].style.display = "none";
		}
	});

	var tableJan = $('#calJan').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"filter": false,
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Jan", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},

		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
					}
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }
					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }
					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
					}
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }
					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "",
				"title": "Jan",
				"autoWidth": false,
				render: function (data, type, row) {
					renderCJan++;

					if (renderCJan === 2) {
						return row.Week;
					}
					if (renderCJan % 7 === 2) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1", 
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2", 
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3", 
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calJan_paginate')[0].style.display = "none";
		}
	});

	var tableFeb = $('#calFeb').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 29,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Feb", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Feb", "autoWidth": false,
				render: function (data, type, row) {
					renderCFeb++;

					if (renderCFeb % 7 === 6) {
						return row.Week;
					}
				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calFeb_paginate')[0].style.display = "none";
		}
	});

	var tableMar = $('#calMar').dataTable({
		"autoWidth": false,
		"processing": true,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"serverSide": true,
		"filter": false,
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Mar", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Mar", "autoWidth": false,
				render: function (data, type, row) {
					renderCMar++;

					if (renderCMar % 7 === 5) {
						return row.Week;
					}
				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calMar_paginate')[0].style.display = "none";
		}
	});

	var tableApr = $('#calApr').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 30,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Apr", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Apr", "autoWidth": false,
				render: function (data, type, row) {
					renderCApr++;

					if (renderCApr % 7 === 2) {
						return row.Week;
					}
				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calApr_paginate')[0].style.display = "none";
		}

	});

	var tableMay = $('#calMay').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "May", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "May", "autoWidth": false,
				render: function (data, type, row) {
					renderCMay++;

					if (renderCMay % 7 === 0) {
						return row.Week;
					}
				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calMay_paginate')[0].style.display = "none";
		}
	});

	var tableJune = $('#calJun').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 30,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Jun", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Jun", "autoWidth": false,
				render: function (data, type, row) {
					renderCJun++;

					if (renderCJun % 7 === 4) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calJun_paginate')[0].style.display = "none";
		}

	});

	var tableJuly = $('#calJul').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Jul", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Jul", "autoWidth": false,
				render: function (data, type, row) {
					renderCJul++;

					if (renderCJul % 7 === 2) {
						return row.Week;
					}
				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calJul_paginate')[0].style.display = "none";
		}
	});

	var tableAug = $('#calAugust').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Aug", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Aug", "autoWidth": false,
				render: function (data, type, row) {
					renderCAug++;

					if (renderCAug % 7 === 6) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calAugust_paginate')[0].style.display = "none";
		}
	});

	var tableSept = $('#calSept').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 30,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Sept", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Sep", "autoWidth": false,
				render: function (data, type, row) {
					renderCSept++;

					if (renderCSept % 7 === 3) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calSept_paginate')[0].style.display = "none";
		}

	});

	var tableOct = $('#calOkt').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Oct", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Oct", "autoWidth": false,
				render: function (data, type, row) {
					renderCOct++;

					if (renderCOct % 7 === 1) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calOkt_paginate')[0].style.display = "none";
		}
	});

	var tableNov = $('#calNov').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 30,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Nov", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Nov", "autoWidth": false,
				render: function (data, type, row) {
					renderCNov++;

					if (renderCNov % 7 === 5) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calNov_paginate')[0].style.display = "none";
		}
	});

	var tableDec = $('#calDec').dataTable({
		"autoWidth": false,
		"processing": true,
		"serverSide": true,
		"filter": false,
		"bLengthChange": false,
		"targets": 'no-sort',
		"bSort": false,
		"order": [],
		"info": false,
		"orderMulti": false,
		"pageLength": 31,
		"destroy": true,
		"ajax": {
			"url": fastApp.BaseUrl + "/Calendar/GetAllCalendarWithParam",
			"data": { strmonth: "Dec", locID: idLoc, year, gtypeID },
			"type": "POST",
			"datatype": "json"
		},
		"columnDefs":
			[{ "targets": [0], "searchable": false, "visible": true, "orderable": false, "defaultContent": null },
			{ 
				"targets": [1], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{ 
				"targets": [2], "searchable": true, "orderable": false, "width": "100%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (rowData.Shift1.indexOf("#") >= 0) {
						$(td).css('background-color', rowData.Shift1);
					}
				}
			},
			{
				"targets": [3], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [4], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}
			},
			{
				"targets": [5], "searchable": true, "orderable": false, "width": "90%",
				"createdCell": function (td, cellData, rowData, row, col) {
					if (cellData === 'A') {
						$(td).css('background-color', '#ff99cc');
					}
					if (cellData === 'B') {
						$(td).css('background-color', '#ffcc99');
					}
					if (cellData === 'C') {
						$(td).css('background-color', '#ffff99');
                    }
                    if (cellData === 'D') {
                        $(td).css('background-color', '#e2efdb');
                    }
                    if (cellData === '-') {
                        $(td).css('background-color', '#833c0c');
                    }

					if (cellData.indexOf("#") >= 0) {
						$(td).css('background-color', cellData);
					}
				}

			},
			{ "targets": [6], "searchable": true, "orderable": true, "visible": false },
			{ "targets": [7], "searchable": true, "orderable": true, "visible": false }
			],

		"columns": [
			{
				"data": "ID", "defaultContent": "", "title": "Dec", "autoWidth": false,
				render: function (data, type, row) {
					renderCDec++;

					if (renderCDec % 7 === 3) {
						return row.Week;
					}

				}
			},
			{
				"data": "Date",
				"title": "Day",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("dd");
					else
						return '';
				}
			},
			{
				"data": "Date",
				"title": "Date",
				render: function (data, type, row) {
					if (data)
						return moment(data).format("D");
					else
						return '';
				}
			},
			{
				"data": "Shift1",
				"title": "Shift 1",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift2",
				"title": "Shift 2",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{
				"data": "Shift3",
				"title": "Shift 3",
				render: function (data, type, row) {
					if (data.indexOf("#") >= 0)
						return '';
					else
						return data;
				}
			},
			{ "data": "Location", "name": "Location", "autoWidth": true },
			{ "data": "GroupType", "name": "GroupType", "autoWidth": true }
		],
		"fnDrawCallback": function () {
			$('#calDec_paginate')[0].style.display = "none";
		}
	});
}