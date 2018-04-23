<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <link rel="stylesheet" href="styles/bootstrap.min.css">
    <link rel="stylesheet" href="styles/notify.css">
    <link rel="stylesheet" href="styles/prettify.css">
    <link rel="stylesheet" href="styles/fontawesome-all.min.css">
    <script src="scripts/jquery.min.js"></script>
    <script src="scripts/popper.min.js"></script>
    <script src="scripts/bootstrap.min.js"></script>
    <script src="scripts/notify.js"></script>
    <script src="scripts/prettify.js"></script>
    <title>bbTools</title>
    <script>
        function ajaxfileupload() {
            var file_data = $('#fileToUpload').prop('files')[0];
            var form_data = new FormData();
            form_data.append('fileToUpload', file_data);
            $.ajax({
                url: 'costcodebackend.php', // point to server-side PHP script
                dataType: 'text',  // what to expect back from the PHP script, if anything
                cache: false,
                contentType: false,
                processData: false,
                data: form_data,
                type: 'post',
                success: function(data){
                    if(data=="error") {
                        $.notify("<span class='fas fa-exclamation-circle'></span> ERROR: Please contact <a href=\"https://github.com/jasonkwh\">Jason Huang</a>.", {type:"danger",close:true,delay:3000});
                    } else {
                        location.href = data;
                    }
                }
            });
        }

        function ajaxmonthlysalary() {
            var file_data = $('#activityupload').prop('files')[0];
            var file_data2 = $('#monthlyupload').prop('files')[0];
            var form_data = new FormData();
            form_data.append('activityupload', file_data);
            form_data.append('monthlyupload', file_data2);
            $.ajax({
                url: 'monthlybackend.php', // point to server-side PHP script
                dataType: 'text',  // what to expect back from the PHP script, if anything
                cache: false,
                contentType: false,
                processData: false,
                data: form_data,
                type: 'post',
                success: function(data){
                    console.log(data);
                    /*if(data=="error") {
                        $.notify("<span class='fas fa-exclamation-circle'></span> ERROR: Please contact <a href=\"https://github.com/jasonkwh\">Jason Huang</a>.", {type:"danger",close:true,delay:3000});
                    } else {
                        location.href = data;
                    }*/
                }
            });
        }
    </script>
</head>
<body>
<div class="container">
    <div class="row justify-content-center" style="margin-top:20px">
        <ul class="nav nav-pills mb-3" id="pills-tab" role="tablist">
            <li class="nav-item">
                <a class="nav-link active" id="pills-srd-tab" data-toggle="pill" href="#pills-srd" role="tab" aria-controls="pills-srd" aria-selected="true">SRD Fortnightly</a>
            </li>
            <li class="nav-item">
                <a class="nav-link" id="pills-mth-tab" data-toggle="pill" href="#pills-mth" role="tab" aria-controls="pills-mth" aria-selected="false">Monthly Salary</a>
            </li>
        </ul>
    </div>
    <div class="row justify-content-center">
        <div class="tab-content" id="pills-tabContent">
            <div class="tab-pane fade show active" id="pills-srd" role="tabpanel" aria-labelledby="pills-srd-tab">
                <div class="card" style="width: 30rem;">
                    <div class="card-body">
                        <h5 class="card-title">Open an Excel file</h5>
                        <div class="row">
                            <div class="col-8">
                                <p class="card-text"><input id="fileToUpload" name="fileToUpload" type="file"></p>
                            </div>
                            <div class="col-4">
                                <button type="button" class="btn btn-success" style="margin-top:-10px" onclick="ajaxfileupload()"><i class="far fa-file-excel"></i>&nbsp;&nbsp;Submit</button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="tab-pane fade" id="pills-mth" role="tabpanel" aria-labelledby="pills-mth-tab">
                <div class="card" style="width: 30rem;">
                    <div class="card-body">
                        <h5 class="card-title">Activity Summary</h5>
                        <div class="row" style="margin-left:1px;margin-top:-8px;margin-bottom:10px">
                            <p class="card-text"><input id="activityupload" name="activityupload" type="file"></p>
                        </div>
                        <h5 class="card-title">GWA Monthly</h5>
                        <div class="row" style="margin-left:1px;margin-top:-8px;margin-bottom:10px">
                            <p class="card-text"><input id="monthlyupload" name="monthlyupload" type="file"></p>
                        </div>
                        <div class="row justify-content-center">
                            <button type="button" class="btn btn-success" onclick="ajaxmonthlysalary()"><i class="far fa-file-excel"></i>&nbsp;&nbsp;Submit</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class="row justify-content-center" style="margin:10px"><p>made with <i class="fab fa-gratipay" style="color:red"></i> by <a href="https://github.com/jasonkwh">Jason Huang</a> <?php echo date("Y"); ?>. Love BB forever!</p></div>
</div>
</body>
</html>