App.directive('excelDownload', excelDownload);

function excelDownload(HttpService,$rootScope) {
    var directive = {
        restrict: 'E',
        transclude: true,
        scope: {},
        template: '<button class="btn btn-primary" onclick="event.preventDefault()">Excel Downlad</button>',
        controller: excelController,
        controllerAs: "excel",
        link: link
    };

    function link(scope, element, attrs, ctrl) {
        var section = element.attr("section");

        element.bind('click', function() {
            ctrl.JSONToCSV(section);
        });
    }

    function excelController() {
        var vm = this;
        vm.data = [];

        vm.exportFromJson = function(data,fileName) {
            //If Jata is not an object then JSON.parse will parse the JSON string in an Object
            var arrData = typeof data != 'object' ? JSON.parse(data) : data;

            var CSV = '';
            CSV += '<tr>';

            //This loop will extract the label from 1st index of on array
            for (var index in arrData[0]) {

                //Now convert each value to string
                CSV += '<td>'+ index +'</td>';
            }

            CSV += '</tr>';

            //1st loop is to extract each row
            for (var i = 0; i < arrData.length; i++) {
                CSV += '<tr>';

                //2nd loop will extract each column and convert it in string
                for (var index in arrData[i]) {
                    CSV += '<td>'+ arrData[i][index] +'</td>';
                }

                CSV += '</tr>';
            }

            var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>'+ CSV +'</table></body></html>'
            var uri = 'data:application/vnd.ms-excel;base64,' + window.btoa(unescape(encodeURIComponent(template)));

            //this trick will generate a temp <a /> tag
            var link = document.createElement("a");
            link.href = uri;

            //set the visibility hidden so it will not effect on your web-layout
            link.style = "visibility:hidden";
            link.download = fileName + ".xls";

            //this part will append the anchor tag and remove it after automatic click
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        vm.JSONToCSV = function(section) {
            var JSONData = [];
            var data;

            if(section == "hotel" && $rootScope.params){
                var params = $rootScope.params;

                HttpService.postJson("json api post url", params).then(function successCallback(response) {

                        vm.data = response.data.returnObject;
						
						// define column names of excel file
                        var cols = [];
                        angular.forEach(vm.pricelist, function (value,key) {
                            cols[key] = {};
                            cols[key]["col 1 name"] = value.col1;
                            cols[key]["col 2 name"] = value.col2;
                        });

                        JSONData.push(cols);

                        angular.forEach(JSONData, function (value,key) {
                            data = value;
                        });

                        vm.exportFromJson(data,"Name of the file");
                        $rootScope.params = {};
                    },function errorCallback(response) {
                        console.log('error');
                    }
                );

            }
        }

    }

    return directive;
}
