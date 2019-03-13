$('#upload').on('click', function() {
    var excelFile = $('#fileUpload')[0];
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    if (regex.test(excelFile.value.toLowerCase())) {
        if (typeof(FileReader) != 'undefined') {
            var reader = new FileReader();
            if (reader.readAsBinaryString) {  // For browsers other than IE.
                reader.onload = function (e) {
                    processExcel(e.target.result);
                };
                reader.readAsBinaryString(excelFile.files[0]);
            } else {  // For IE.
                $('#dvExcel').append('2333');
            }
        } else {
            alert('The browser does not support HTML5.  Please use Google Chrome.');
        }
    } else {
        alert('Please provide a valid MS Excel file.');
    }
});

function processExcel(data) {
    var workbook = XLSX.read(data, { type: 'binary' });
    var sheetName = workbook.SheetNames[0];  // Target sheet has to be the first sheet in the workbook.
    var sheet = workbook.Sheets[sheetName];

    var targetColumns = ['A', 'C', 'I', 'K', 'L'];
    var students = [];
    var row = 3;
    while (sheet[targetColumns[0] + row.toString()]) {
        var studentName = sheet[targetColumns[0] + row.toString()].v,
            studentID = parseInt(sheet[targetColumns[1] + row.toString()].v),
            outcomeName = sheet[targetColumns[2] + row.toString()].v,
            attempt = parseInt(sheet[targetColumns[3] + row.toString()].v),
            score = parseInt(sheet[targetColumns[4] + row.toString()].v);
        var index = findStudent(students, studentID);
        if (index === -1) {
            students.push(new Student(studentName, studentID));
            students[students.length - 1].scores[outcomeName] = [attempt, score];
        } else {
            var currentStudent = students[index];
            if (!(outcomeName in currentStudent.scores)) {
                currentStudent.scores[outcomeName] = [attempt, score];
            } else {
                currentStudent.scores[outcomeName] = [currentStudent.scores[outcomeName][0] + 1, currentStudent.scores[outcomeName][1] + score];
            }
        }
        row++;
    }
    outputStudents(students, '#output');
}

function findStudent(students, id) {
    for (var i = 0; i < students.length; i++) {
        if (id === students[i].id) {
            return i;
        }
    }
    return -1;
}

function outputStudents(students, divID) {
    var CLOs = [];
    var table = '<table style="border-collapse:collapse;border:none;width:100%;text-align: center;">';
    table += '<tbody><tr>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Student Name</td>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Student ID</td>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Learning Outcome Name</td>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Attempts</td>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Score</td>';
    table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">Percentage</td>';
    table += '</tr>';
    for (var i = 0; i < students.length; i++) {
        var rowSpan = Object.keys(students[i].scores).length;
        for (var j = 0; j < Object.keys(students[i].scores).length; j++) {
            table += '<tr>';
            if (j === 0) {
                table += '<td style = "border:none;border: 2px solid blue;text-align:left;padding:2pt 4pt;" rowspan="' + rowSpan.toString() + '">' + students[i].name + '</td>';
                table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;" rowspan="' + rowSpan.toString() + '">' + students[i].id.toString() + '</td>';
            }
            table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;">' + Object.keys(students[i].scores)[j] + '</td>';
            table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;">' + students[i].scores[Object.keys(students[i].scores)[j]][0].toString() + '</td>';
            table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;">' + students[i].scores[Object.keys(students[i].scores)[j]][1].toString() + '</td>';
            var percent = students[i].scores[Object.keys(students[i].scores)[j]][1] * 1.0 / students[i].scores[Object.keys(students[i].scores)[j]][0];
            percent = Math.round(percent * 100);
            var index = findCLO(CLOs, Object.keys(students[i].scores)[j]);
            if (index === -1) {
                CLOs.push(new CLO(Object.keys(students[i].scores)[j]));
                index = CLOs.length - 1;
            }
            if (percent >= 90) { CLOs[index].numOfMastery++; }
            else if (percent >= 70) { CLOs[index].numOfCompetence++; }
            else { CLOs[index].numOfLowSkill++; }
            table += '<td style = "border:none;border: 2px solid blue;padding:2pt 4pt;font-weight:bold;">' + percent.toString() + '%</td>';
            table += '</tr>';
        }
    }
    table += '</tbody></table>';
    $(divID).html(table);
    google.charts.load('current', {'packages':['bar']});
    google.charts.setOnLoadCallback(drawChart);
    function drawChart() {
        var dataArr = [['CLO', 'Number of Mastery', 'Number of Competence', 'Number of Low Skill']];
        for (var i = 0; i < CLOs.length; i++) {
            dataArr.push([CLOs[i].name, CLOs[i].numOfMastery, CLOs[i].numOfCompetence, CLOs[i].numOfLowSkill]);
        }
        var data = google.visualization.arrayToDataTable(dataArr);
        var options = {
            backgroundColor: 'transparent',
            fontName: 'Tahoma',
            fontSize: '15',
            height: '300',
            chart: { title: 'CLO Summary' }
        };
        var chart = new google.charts.Bar(document.getElementById('chart'));
        chart.draw(data, google.charts.Bar.convertOptions(options));
    }
}

function findCLO(container, name) {
    for (var i = 0; i < container.length; i++) {
        if (name === container[i].name) {
            return i;
        }
    }
    return -1;
}