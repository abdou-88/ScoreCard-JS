
////// variables declaration area //////
var loginName ="";
var Category = "1";
var CurrentAnaly = "";
var SelectedMonth = {
  mm: new Date().getMonth() + 1,
  yyyy: new Date().getFullYear(),
}; 
var MonthFound;
var index;
var ChartRendered = false; 

// Scorecard Data

var ScoreCardData = [];

// productivity chart options and data
var ProdOptions = {
  series: [
    {
      name: "Productivity",
      data: [],
    },
  ],
  chart: {
    id: "chart1",
    height: 350,
    width: "100%",
    type: "bar",
    toolbar: {
      show: false,
    },
  },
  annotations: {
    yaxis: [
      {
        y: "95",
        borderColor: "red",
        label: {
          borderColor: "#00E396",
          orientation: "vertical",
        },
      },
    ],
  },
  plotOptions: {
    bar: {
      horizontal: false,
      columnWidth: "60%",
    },
  },
  stroke: {
    show: true,
    width: 2,
  },
  dataLabels: {
    enabled: false,
  },
  yaxis: {
    labels: {
      formatter: function (val) {
        return val + "%";
      },
    },
  },
  fill: {
    opacity: 1,
    colors: "#0C9ED9",
  },
  xaxis: {
    type: "text",
  },
};
// Availability chart options and data
var AvOptions = {
  series: [
    {
      name: "Availability",
      data: [],
    },
  ],
  chart: {
    id: "chart2",
    height: 350,
    width: "100%",
    type: "bar",
    toolbar: {
      show: false,
    },
  },
  annotations: {
    yaxis: [
      {
        y: "80",
        borderColor: "red",
        label: {
          borderColor: "#00E396",
          orientation: "vertical",
        },
      },
    ],
  },
  plotOptions: {
    bar: {
      horizontal: false,
      columnWidth: "60%",
    },
  },
  stroke: {
    show: true,
    width: 2,
  },
  dataLabels: {
    enabled: false,
  },
  yaxis: {
    labels: {
      formatter: function (val) {
        return val + "%";
      },
    },
  },
  fill: {
    opacity: 1,
    colors: "#0C9ED9",
  },
  xaxis: {
    type: "text",
  },
};


////////////// Functions area \\\\\\\\\\\\\\

//Getting-retreving scord card data
function retrieveMonthScoreCardData(TheMonth) {
  // show the loading over layer
  document.getElementById("overlayLoading").style.display = "block";
 
  SP.SOD.executeFunc(
    "SP.js",
    "SP.ClientContext",

    function () {
      var clientContext = new SP.ClientContext();
      var ScheList = clientContext.get_web().get_lists().getByTitle("ScoreCard");
      var camlQuery = new SP.CamlQuery();

      ScoreCardListItem = ScheList.getItems(camlQuery);
      clientContext.load(ScoreCardListItem);

      clientContext.executeQueryAsync(
        Function.createDelegate(this, onRetrieveSucceeded),
        Function.createDelegate(this, onRetrieveFailed)
      );


      function onRetrieveSucceeded() {

        // getting the data of the ScoreCard
        var listItemEnumerator = ScoreCardListItem.getEnumerator();
        MonthFound = "0";
        
        while (listItemEnumerator.moveNext() && MonthFound == "0") {
          var oListItem = listItemEnumerator.get_current();
          
          if (oListItem.get_item("Month") == TheMonth) {
            ScoreCardData = JSON.parse($(oListItem.get_item("ScoreCardData")).text());

            MonthFound = "1";
          }
        }

        // here we show the data of the first analyst
       
        GetCurrentUser();
        
        IsAdmin("SCAdmins")
          ? (index = 0)
          : (index = ScoreCardData.findIndex(
              (SCard) => SCard.AnalystName === loginName
            ));
        
        if (index == -1){
          alert("There is no ScoreCard data for you this month, please contact your TL!");
          document.getElementById("overlayLoading").style.display =
            "none";
        }

        //loading data on the page
        ShowSCData(ScoreCardData[index]);
        ShowPopUpData(ScoreCardData[index], Category);
        LoadCharts();
        // remove the loading OverLayer
        document.getElementById("overlayLoading").style.display = "none";
       
      }

      function onRetrieveFailed(sender, args) {
        alert(
          "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
        );
      }

    }
  );
};

// getting the current sharepoint username 
function GetCurrentUser() {

    var userid = _spPageContextInfo.userId;
    var requestUri =_spPageContextInfo.webAbsoluteUrl +"/_api/web/getuserbyid(" +userid +")";
    var requestHeaders = { accept: "application/json;odata=verbose" };

    $.ajax({
    url: requestUri,
    contentType: "application/json;odata=verbose",
    headers: requestHeaders,
    success: onSuccess,
    error: onError,
    });

    function onSuccess(data, request) {

      loginName = data.d.Title;
      
      var select = document.getElementById("AName");
      
      if (IsAdmin("SCAdmins")) {
        // if it is admin then we fill out the dropdown with all analysts
        select.options.length = 0;
        ScoreCardData.forEach(
          (SCard) =>
            (select.options[select.options.length] = new Option(
              SCard.AnalystName,
              SCard.AnalystName
            ))
        );
        
      } else {
        //if user is not an admin we just show his name 
        select.options[select.options.length] = new Option(loginName, loginName);
      }
        
    }

    function onError(error) {
    alert(error);
    }

};

//checking the permissions of this user
function IsAdmin(groupName) {
  var userIsInGroup = false;
  $.ajax({
    async: false,
    headers: { accept: "application/json; odata=verbose" },
    method: "GET",
    url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/currentuser/groups",
    success: function (data) {
      data.d.results.forEach(function (value) {
        if (value.Title === groupName) {
          userIsInGroup = true;
        }
      });
      
    },
    error: function (response) {

      console.log(response.status);
    },
    
  });
  
  return userIsInGroup;
}

// Here we show the Scorecard data on the page
function ShowSCData(AnData){

    // Productivity
     var ProdAvrg =
       (parseInt(AnData.Scorecard.productivity.Calls.percentage) +
         parseInt(AnData.Scorecard.productivity.Emails.percentage)) /
       2; 
    document.getElementById("Prod").textContent = ProdAvrg + "%";

    ProdAvrg > 95
      ? document.getElementById("Prod").setAttribute("fill", "green")
      : document.getElementById("Prod").setAttribute("fill", "red");
     
    // Availability
    document.getElementById("Avail").textContent =
      AnData.Scorecard.Availability.percentage + "%";

    AnData.Scorecard.Availability.percentage > 80
      ? document.getElementById("Avail").setAttribute("fill", "green")
      : document.getElementById("Avail").setAttribute("fill", "red");
    
    // Quality  
    var qualAvrg =
      (parseInt(AnData.Scorecard.Quality.Calls) +
        parseInt(AnData.Scorecard.Quality.Tickets)) /
      2; 

    document.getElementById("Qual").textContent = qualAvrg + "%";

    qualAvrg < 80
      ? document.getElementById("Qual").setAttribute("fill", "red")
      : qualAvrg < 84.9
      ? document.getElementById("Qual").setAttribute("fill", "orange")
      : qualAvrg < 91.9
      ? document.getElementById("Qual").setAttribute("fill", "yellow")
      : qualAvrg < 96.9
      ? document.getElementById("Qual").setAttribute("fill", "#9ACD32")
      : document.getElementById("Qual").setAttribute("fill", "green");

    // the rest of categories
    document.getElementById("Mgoal").textContent =
        AnData.Scorecard.MonthGoal;
    
    document.getElementById("csat").textContent =
        AnData.Scorecard.Csat;

    document.getElementById("Cmnt").textContent =
        AnData.Scorecard.Comment;

 
};

// Here we show the more info popup data on the page depends on the category
function ShowPopUpData(AnData,Category) {
  switch (Category) {
    case "1": {
      $("#MITitle").html("Productivity");
      $("#MIRight1").html("<h2>Emails Prod:</h2>" +AnData.Scorecard.productivity.Emails.percentage +"%");
      $("#MILeft1").html("<h2>Calls Prod:</h2>" +AnData.Scorecard.productivity.Calls.percentage + "%");
      $("#MIRight2").html("<h2>Emails count:</h2>" +AnData.Scorecard.productivity.Emails.Count );
      $("#MILeft2").html("<h2>Calls count:</h2>" +AnData.Scorecard.productivity.Calls.Count);
      $("#MIRight3").html("");
      $("#MILeft3").html("" );

      break;
    }
    case "2": {
      $("#MITitle").html("Availability");
      $("#MIRight1").html("<h2>Availability:</h2>" +AnData.Scorecard.Availability.percentage +"%");
      $("#MILeft1").html("<h2>Rona :</h2>" +AnData.Scorecard.Availability.Rona );
      $("#MIRight2").html("");
      $("#MILeft2").html("");
      $("#MIRight3").html("");
      $("#MILeft3").html("");
      
      break;
    }
    case "3": {
      $("#MITitle").html("Quality");
      $("#MIRight1").html("<h2>Calls Quality:</h2>" + AnData.Scorecard.Quality.Calls + "%");
      $("#MILeft1").html("<h2>Tickets Quality:</h2>" + AnData.Scorecard.Quality.Tickets + "%");
      $("#MIRight2").html("");
      $("#MILeft2").html("");
      $("#MIRight3").html("");
      $("#MILeft3").html("");

      break;
    }
    default: {
      console.log("Invalid choice" + Category);
      break;
    }
  }
};

// More info function changing Category and showing popup
function MoreInfo(Catg) {
  Category = Catg;
  ChangeAnalyst();
  $(".PopDivClass").show();
};

// kanloadiw this 2 the charts on the right
function LoadCharts(){

  ProdOptions.series[0].data = [];
  AvOptions.series[0].data = [];

  // filling the chart data of productivity and availability depends on the permissions
  if (IsAdmin("SCAdmins")) {
    // giving the admin permissions
    ScoreCardData.forEach((Scard) => {
      var ProdChart = {
        x: Scard.AnalystName,
        y:(parseInt(Scard.Scorecard.productivity.Calls.percentage) +parseInt(Scard.Scorecard.productivity.Emails.percentage)) / 2,
      };
      var AVChart = {
        x: Scard.AnalystName,
        y: Scard.Scorecard.Availability.percentage,
      };

      ProdOptions.series[0].data.push(ProdChart);
      AvOptions.series[0].data.push(AVChart);
    });
  } else {
    //giving a user permissions 
    ScoreCardData.forEach((Scard) => {
      if (Scard.AnalystName == loginName) {
        var ProdChart = {
          x: Scard.AnalystName,
          y: (parseInt(Scard.Scorecard.productivity.Calls.percentage) +parseInt(Scard.Scorecard.productivity.Emails.percentage)) / 2,
          fillColor: "#000000",
          strokeColor: "#000000",
        };
        var AVChart = {
          x: Scard.AnalystName,
          y: Scard.Scorecard.Availability.percentage,
          fillColor: "#000000",
          strokeColor: "#000000",
        };
      } else {
        var ProdChart = {
          x: "",
          y:(parseInt(Scard.Scorecard.productivity.Calls.percentage) +parseInt(Scard.Scorecard.productivity.Emails.percentage)) / 2,
        };
        var AVChart = {
          x: "",
          y: Scard.Scorecard.Availability.percentage,
        };
      }

      ProdOptions.series[0].data.push(ProdChart);
      AvOptions.series[0].data.push(AVChart);
    });
  }

  // sort out the avail and prod charts
  AvOptions.series[0].data = AvOptions.series[0].data.sort(function (a, b) {return a.y - b.y;});
  ProdOptions.series[0].data = ProdOptions.series[0].data.sort(function (a, b) {return a.y - b.y;});

  // initial the productivity and avail charts
  var ProdChart = new ApexCharts(document.querySelector("#Prodchart"),ProdOptions);
  var AvChart = new ApexCharts(document.querySelector("#Avchart"), AvOptions);

  if (!ChartRendered) {
    // we render the chart for first time
    AvChart.render();
    ProdChart.render();
    ChartRendered = true;

    if (IsAdmin("SCAdmins")){
      //also we create the admins buttons for the first time
      
      document.getElementById("EditBTN").style.display="block";

      var LoadExcelbtn = document.createElement("BUTTON");
      LoadExcelbtn.innerHTML = "Load Data";
      LoadExcelbtn.style.display = "none";

      var cmntbtn = document.createElement("BUTTON");
      cmntbtn.innerHTML = "Edit comment";
      cmntbtn.style.display = "none";

      var PubSCbtn = document.createElement("BUTTON");
      PubSCbtn.innerHTML = "Publish ScoreCard";
      PubSCbtn.onclick = function () {
        SaveScoreCard(ScoreCardData, NavigateMonths(0), "ScoreCard");
      };
      PubSCbtn.style.display = "none";
    // setting the class, type and id of the btns
    LoadExcelbtn.className = "NavButton";
    cmntbtn.className = "NavButton";
    PubSCbtn.className = "NavButton";

    LoadExcelbtn.id = "LoadExcelData";
    cmntbtn.id = "EditCmnt";
    PubSCbtn.id = "PubSc";

    LoadExcelbtn.type = "button";
    cmntbtn.type = "button";
    PubSCbtn.type = "button";

    //adding it to the div area
    document.getElementById("AdminBtnsArea").appendChild(LoadExcelbtn);
    document.getElementById("AdminBtnsArea").appendChild(cmntbtn);
    document.getElementById("AdminBtnsArea").appendChild(PubSCbtn);

    }
      

  } else {
    // here we update the data of both charts for each month :D
    ApexCharts.exec("chart1", "updateSeries", [
      {
        data: ProdOptions.series[0].data,
      },
    ]);
    ApexCharts.exec("chart2", "updateSeries", [
      {
        data: AvOptions.series[0].data,
      },
    ]);
  }

};

// here we handle the change of the analyst name on the drop down
function ChangeAnalyst(){

  var an = document.getElementById("AName");
  var AnalystName = an.options[an.selectedIndex].value;
  
  index = ScoreCardData.findIndex((SCard) => SCard.AnalystName === AnalystName);
      
  ShowSCData(ScoreCardData[index]);
  ShowPopUpData(ScoreCardData[index], Category);

};

// to form the date and navigate the months
function NavigateMonths(NumberM) {
  var MonthAdded = false;
  var ShowMonthString;
  var ReturnMonthString;
  
  if (SelectedMonth.mm == 12 && NumberM != -1) {
    SelectedMonth.mm = 1;
    SelectedMonth.yyyy = SelectedMonth.yyyy + 1;
    MonthAdded = true;
  } else {
    if (SelectedMonth.mm == 1 && NumberM == -1) {
      SelectedMonth.mm = 12;
      SelectedMonth.yyyy = SelectedMonth.yyyy - 1;
      MonthAdded = true;
    }
  }
    
  if (SelectedMonth.mm < 10 ) {
    if (!MonthAdded) {
      SelectedMonth.mm = SelectedMonth.mm + NumberM;
    }
    mm = "0" + SelectedMonth.mm;
    ReturnMonthString = SelectedMonth.yyyy + "-" + mm;
    ShowMonthString = MonthToString(mm) + " " + SelectedMonth.yyyy;
  } else {
    if (!MonthAdded) {
      SelectedMonth.mm = SelectedMonth.mm + NumberM;
    }
    ReturnMonthString = SelectedMonth.yyyy + "-" + SelectedMonth.mm;
    ShowMonthString = MonthToString(SelectedMonth.mm) + " " + SelectedMonth.yyyy;
  };

  
  document.getElementById("MonthHeader").innerHTML = ShowMonthString;
  
  return ReturnMonthString;
};

function MonthToString(MNum){
  switch (MNum) {
    case "01":
      return "January";
    case "02":
      return "February";
    case "03":
      return "March";
    case "04":
      return "April";
    case "05":
      return "May";
    case "06":
      return "June";
    case "07":
      return "July";
    case "08":
      return "August";
    case "09":
      return "September";
    case "10":
      return "October";
    case "11":
      return "November";
    case "12":
      return "December";
    default:
      return null;
  }

};



//save the ScoreCard data
function SaveScoreCard(SCData, themonth, Listname) {

  document.getElementById("overlayLoading").style.display = "block";
    
  var clientContext = SP.ClientContext.get_current();
  list = clientContext.get_web().get_lists().getByTitle(Listname);
  var skillcamlQuery = new SP.CamlQuery();

  skillcamlQuery.set_viewXml(
    "<View><Query><Where><Eq><FieldRef Name='Month' /><Value Type='Text'>" +
      themonth +
      "</Value></Eq></Where></Query></View>"
  );

  this.skillcollListItem = list.getItems(skillcamlQuery);

  clientContext.load(skillcollListItem);

  clientContext.executeQueryAsync(
    Function.createDelegate(this, function () {_returnParam = onSavingSucceeded();}),
    Function.createDelegate(this, function () {_returnParam = onSavingFailed();})
  );

  
  function onSavingSucceeded() {

    const ScoreCardData = JSON.stringify(SCData);
    
    // if the month already exist - we update
    if (skillcollListItem.get_count() >= 1) {
      var listItemEnumerator = skillcollListItem.getEnumerator();

      while (listItemEnumerator.moveNext()) {
        var oListItem = listItemEnumerator.get_current();

        if (oListItem.get_item("Month") == themonth) {
          oListItem.set_item("ScoreCardData", ScoreCardData);
          
          oListItem.update();
        }
      }
    }
    // if it is a new month - we add new item
    else {
      
      var listItemCreationInfo = new SP.ListItemCreationInformation();
      var newItem = list.addItem(listItemCreationInfo);

      newItem.set_item("Month", themonth);
      newItem.set_item("ScoreCardData", SCData);
      

      newItem.update();
    }

    clientContext.executeQueryAsync(
      Function.createDelegate(this, onSavedSuccess),
      Function.createDelegate(this, onSavedFailure)
    );



    function onSavedSuccess() {
        document.getElementById("overlayLoading").style.display = "none";
        savedBol = "0";
        //alertify.set({ delay: 5000 });
        //alertify.success("The ScoreCard has been published successfully");
        console.log("Saved");

    }

    function onSavedFailure(args) {

        //alertify.set({ delay: 5000 });
        //alertify.error("Something went wrong!!");
        
        return false;
    }

  }

  function onSavingFailed(sender, args) {
    alert(
      "Request failed. " + args.get_message() + "\n" + args.get_stackTrace()
    );
    return false;
  }

};







// Jquery ready page function 
$(document).ready(function () {
  

  retrieveMonthScoreCardData(NavigateMonths(-2));

  $("#previous").click(function () {
    retrieveMonthScoreCardData(NavigateMonths(-1));
  });
  $("#next").click(function () {
    retrieveMonthScoreCardData(NavigateMonths(1));
  });

  $("#EditCheck").click(function (){
    if ($(this).is(":checked")) {
      document.getElementById("LoadExcelData").style.display = "block";
      document.getElementById("EditCmnt").style.display = "block";
      document.getElementById("PubSc").style.display = "block";
      
      console.log("checked")
    } else if ($(this).is(":not(:checked)")) {
      document.getElementById("LoadExcelData").style.display = "none";
      document.getElementById("EditCmnt").style.display = "none";
      document.getElementById("PubSc").style.display = "none";
    }
  });
  $("#LoadExcelData").click(function () {});

  $("#EditCmnt").click(function () {});



  $("#AName").on("change", function () {
    ChangeAnalyst();
  });

  //popup closing js code
  $(".PopDivClass").click(function () {
    $(".PopDivClass").hide();
  });
  $(".popupCloseButton").click(function () {
    $(".PopDivClass").hide();
  });
  /////////
})

