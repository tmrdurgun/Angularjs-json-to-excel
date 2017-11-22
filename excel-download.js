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
            //If JSONData is not an object then JSON.parse will parse the JSON string in an Object
            var arrData = typeof data != 'object' ? JSON.parse(data) : data;

            var CSV = '';
            //Set Report title in first row or line

            CSV += '\r\n\n';

            var row = "";

            //This loop will extract the label from 1st index of on array
            for (var index in arrData[0]) {

                //Now convert each value to string and comma-seprated
                row += index + ',';
            }

            row = row.slice(0, -1);

            //append Label row with line break
            CSV += row + '\r\n';

            //1st loop is to extract each row
            for (var i = 0; i < arrData.length; i++) {
                var row = "";

                //2nd loop will extract each column and convert it in string comma-seprated
                for (var index in arrData[i]) {
                    row += '"' + arrData[i][index] + '",';
                }

                row.slice(0, row.length - 1);

                //add a line break after each row
                CSV += row + '\r\n';
            }

            if (CSV == '') {
                alert("Invalid data");
                return;
            }

            //Initialize file format you want csv or xls
            var uri = 'data:application/csv;charset=utf-8,' + encodeURIComponent(CSV);

            // Now the little tricky part.
            // you can use either>> window.open(uri);
            // but this will not work in some browsers
            // or you will not get the correct file extension

            //this trick will generate a temp <a /> tag
            var link = document.createElement("a");
            link.href = uri;

            //set the visibility hidden so it will not effect on your web-layout
            link.style = "visibility:hidden";
            link.download = fileName + ".csv";

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