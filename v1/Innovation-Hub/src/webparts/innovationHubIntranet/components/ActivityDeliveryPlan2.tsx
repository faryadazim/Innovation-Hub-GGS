

import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import './ActivityPlanStyle2.css'
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  Persona,
  PersonaPresence,
  PersonaSize,
  DatePicker,
  Spinner,
  PrimaryButton,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  TextField,
  Checkbox,
  Modal,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { maxBy } from "lodash";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
const saveIcon: IIconProps = { iconName: "Save" };
const editIcon: IIconProps = { iconName: "Edit" };
const cancelIcon: IIconProps = { iconName: "Cancel" };

let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

const ActivityDeliveryPlan = (props: any) => {
  // Variable-Declaration-Section Starts
  //  const webURL_ = "https://ggsaus.sharepoint.com";
  //  const WeblistURL_ = "Annual Plan";


  // const sharepointWeb = Web(webURL_);
  const sharepointWeb = Web(props.URL);
  const activityPlan_ID = props.ActivityPlanID;

  const activityPlanListName = "Activity Plan";
  const adpListName = "Activity Delivery Plan";
  const templateListName = "Activity Delivery Plan Template";
  const activityPBListName = "ActivityProductionBoard";

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  const adpCurrentWeekNumber = moment().isoWeek();
  const adpCurrentYear = moment().year();

  const allPeoples = props.peopleList;

  const adpStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
  });


  const adpDrpDwnOptns = {
    developerOptns: [{ key: "All", text: "All" }],
    stepsOptns: [{ key: "All", text: "All" }],
    lessonOptns: [{ key: "All", text: "All" }],
    statusOptns: [{ key: "All", text: "All" }],
    weekOptns: [{ key: "All", text: "All" }],
    yearOptns: [{ key: "All", text: "All" }],
  };
  const adpFilterKeys = {
    developer: "All",
    step: "All",
    lesson: "",
    status: "All",
    week: "All",
    year: "All",
  };
  // const adpFilterKeys = { developer: "All", step: "All", lesson: "All" };

  // Variable-Declaration-Section Ends
  // Styles-Section Starts



  const adpCommonStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: 25,
    fontWeight: "600",
    padding: 3,
    width: 100,
    display: "flex",
    justifyContent: "center",
  });

  const adpbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const adpbuttonStyleClass = mergeStyleSets({
    buttonPrimary: [
      {
        color: "White",
        backgroundColor: "#FAA332",
        borderRadius: "3px",
        border: "none",
        marginRight: "10px",
        selectors: {
          ":hover": {
            backgroundColor: "#FAA332",
            opacity: 0.9,
            borderRadius: "3px",
            border: "none",
            marginRight: "10px",
          },
        },
      },
      adpbuttonStyle,
    ],
    buttonSecondary: [
      {
        color: "White",
        backgroundColor: "#038387",
        borderRadius: "3px",
        border: "none",
        margin: "0 5px",
        selectors: {
          ":hover": {
            backgroundColor: "#038387",
            opacity: 0.9,
          },
        },
      },
      adpbuttonStyle,
    ],
  });
  const adpIconStyleClass = mergeStyleSets({
    navArrow: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    navArrowDisabled: [
      {
        cursor: "pointer",
        color: "#ababab",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
    link: [
      {
        fontSize: 17,
        height: 16,
        width: 16,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    linkDisabled: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#ababab",
        cursor: "not-allowed",
        padding: 8,
        borderRadius: 3,
        marginLeft: 10,
      },
    ],
    refresh: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
        color: "#fff",
        backgroundColor: "#038387",
        cursor: "pointer",
        padding: 8,
        borderRadius: 3,
        marginTop: 34,
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    save: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
    edit: [
      {
        fontSize: "18px",
        color: "#fff",
        paddingRight: 10,
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
      },
    ],
  });

  // Styles-Section Ends
  // States-Declaration Starts
  const [Open, setOpen] = useState([])
  const [totalHours, setTotalHours] = useState(0)
  const [adpReRender, setAdpReRender] = useState(true);
  const [currentUser, setCurrentUser] = useState({});
  const [activtyPlanItem, setActivtyPlanItem] = useState([]);
  const [activityPB, setActivityPB] = useState([]);
  const [group, setgroup] = useState([]);
  const [adpMasterData, setAdpMasterData] = useState([]);
  const [adpData, setAdpData] = useState([]);
  const [adpDropDownOptions, setAdpDropDownOptions] = useState(adpDrpDwnOptns);
  const [adpFilters, setAdpFilters] = useState(adpFilterKeys);
  const [adpActivityResponseData, setAdpActivityResponseData] = useState([]);
  const [adpEditFlag, setAdpEditFlag] = useState(false);
  const [newDataFlag, setNewDataFlag] = useState(false);
  const [adpItemAddFlag, setAdpItemAddFlag] = useState(false);
  const [adpLoader, setAdpLoader] = useState("noLoader");
  const [adpLoader2, setAdpLoader2] = useState("noLoader");
  const [adpAutoSave, setAdpAutoSave] = useState(false);
  const [finalStep, setFinalStep] = useState([]);
  const [finalStepConst, setFinalStepConst] = useState([]);
  const [annualPlanData, setAnnualPlanData] = useState({ ba: '', term: 0 });


  const [adpSDSort, setAdpSDSort] = useState("");
  const [adpEDSort, setAdpEDSort] = useState("");

  const [AdpIsCompleted, setAdpIsCompleted] = useState(false);
  const [StepsArray, setStepsArray] = useState([])
  const [reRenderState, setreRenderState] = useState(false)
  const [StepsArrayCustomized, setStepsArrayyCustomized] = useState(
    [
      { name: "Draft", key: "Draft" },
      { name: "Review", key: "Endorsed" },
      { name: "Edit", key: "Edited" },
      { name: "Assemble", key: "Assembled" },
      { name: "Approved ", key: "Signed Off,Publish ready" },
      { name: "Distribute", key: "Approved" }
      // { name: "Distribute", key: "Completed" }
    ]
  )
  const [Tabledata, setTabledata] = useState([])
  const [diliveryPlanNeedToBeUpdate, setDiliveryPlanNeedToBeUpdate] = useState([])
  const [AdpConfirmationPopup, setAdpConfirmationPopup] = useState({
    condition: false,
    isNew: false,
  });
  const formatData = async (records: any) => {
    console.log(records, "form data")
    const newArray = [];
    let currentLesson = records[0]?.LessonID;
    let currentGroup = [{ ...records[0] }];

    for (let i = 1; i < records.length; i++) {
      if (records[i].LessonID === currentLesson) {
        currentGroup.push({ ...records[i] });

      } else {
        newArray.push(currentGroup);
        currentLesson = records[i]?.LessonID;
        currentGroup = [{ ...records[i] }];
      }
    }

    newArray.push(currentGroup);
    const stepsSet = new Set(records.map(obj => obj.Steps));
    const stepsArray = Array.from(stepsSet);
    await getReviewLogInfo(activityPlan_ID, records)
    setTabledata([...newArray])

    setStepsArray(stepsArray)
    hoursCal([...newArray])

  }


  const getReviewLogInfo = (activityPlan_ID: any, records: any) => {
    // AH: 5    Developer: { name: 'Charlie Archbold', id: 236, email: 'carchbold@goodtogreatschools.org.au' }
    // EditorId: 236    End: "14/07/2023"    ID: 36352    IsCompleteNew: false    IsCompleteStatus: false    Lesson: "Lesson 26"    LessonID: 26    MaxPH: ""    MinPH: ""    OrderId: 276    PH: 4    PHError: false    PHWeek: null    Project: "Oz-e-Writing Years F-6 Unit 3
    // Start: "03/07/2023"    Status: "Scheduled"    Steps: "Draft"    Title: "Draft"    Types: "Curriculum (Writing a lesson)"    dateError:
    // false

    let diliveryPlanItems = records;
    console.log(diliveryPlanItems, "inside getReviewLogInfo")
    let joinReviewLogList = []
    sharepointWeb.lists
      .getByTitle("Review Log")
      .items.select("*")
      .filter(`AnnualPlanID eq ${activityPlan_ID}`)
      .top(5000)
      .get()
      .then((items: any) => {

        let reviewLogItems = [];

        let count = 0
        items.map((item) => {


          if (
            item.auditRequestType == "Review" || item.auditRequestType == "Initial Edit" || item.auditRequestType == "Assemble" || item.auditRequestType == "Sign-off" || item.auditRequestType == "Publish" || item.auditRequestType == "Distribute") {

            const diliverySteps = diliveryPlanItems.find(element => element.ID == item.DeliveryPlanID);

            if (diliverySteps == undefined) {
              console.log("no availble", item.DeliveryPlanID)
            } else {
              reviewLogItems.push({
                // FromUserId :item.FromUserId,
                request: item.auditRequestType,
                response: item.auditLastResponse,
                Dev: item.auditFrom,
                FromEmail: item.FromEmail,
                // ToUserId:item.ToUserId,
                client: item.auditTo,
                ToEmail: item.ToEmail,
                Lesson: diliverySteps.Lesson,
                LessonID: diliverySteps.LessonID,
                Project: diliverySteps.Project,
                Start: item.Created,
                // Start: diliverySteps.Start,
                End: item.auditSent,
                Created: item.Created,
                Modified: item.Modified,
                ID: item.ID,

                // End: diliverySteps.End,
                // ReqStart : item
                // ReqEnd : item.auditSent



              })
              count++
            }
            // else if (diliverySteps && diliverySteps.length == 1) {

            // } else {
            //   console.log("Less than one Delivery Plan")
            // }


          }




        })

        function sortByModifiedDateAscending(arr: any[]) {
          // Create a copy of the original array to avoid modifying the original data
          var sortedArray = arr.slice();

          // Sort the array by "Modified" date in ascending order
          sortedArray.sort(function (a: any, b: any) {
            return moment(a.End).diff(moment(b.End));
          });

          return sortedArray;
        }

        // Usage example


        var sortedData = sortByModifiedDateAscending(reviewLogItems);
        // DeliveryPlanID


        // Initialize an empty object to store the grouped lessons
        const groupedLessons = {};

        // Iterate over each object in the revLog array
        sortedData.forEach(item => {
          const { Lesson } = item;

          // Check if the Lesson already exists in the groupedLessons object
          if (groupedLessons.hasOwnProperty(Lesson)) {
            // If it exists, push the current item to the existing lesson array
            groupedLessons[Lesson].lessonsData.push(item);
          } else {
            // If it doesn't exist, create a new lesson array with the current item
            groupedLessons[Lesson] = {
              Lesson,
              lessonsData: [item]
            };
          }
        });

        // Extract the values from the groupedLessons object to get the final array
        const groupedLessonsArray = Object.keys(groupedLessons).map(key => groupedLessons[key]);





        sortedData.sort(function (a, b) {
          return a.LessonID - b.LessonID;
        });



        // First, sort by lesson number in ascending order
        sortedData.sort(function (a, b) {
          return a.LessonID - b.LessonID;
        });

        console.log(sortedData)

        // Then, prioritize rows with responses "Edited," "Feedbacked," or "Endorsed" on top for each LessonID
        // reviewLogItems.sort(function (a, b) {
        //   // var priorityResponses = ['Edited', 'Feedbacked', 'Endorsed'];
        //   var priorityResponses = ['Endorsed', 'Edited', 'Assembled', 'Signed Off', 'Publish ready', 'Completed'];
        //   var responsePriorityA = priorityResponses.indexOf(a.response) !== -1 ? 0 : 1;
        //   var responsePriorityB = priorityResponses.indexOf(b.response) !== -1 ? 0 : 1;

        //   // If both rows have the same LessonID number, prioritize by response value
        //   if (a.LessonID === b.LessonID) {
        //     return responsePriorityA - responsePriorityB;
        //   }

        //   // Otherwise, maintain the lesson number order
        //   return 0;
        // });




        const respo = getUniqueLessonsWithSteps(diliveryPlanItems)
        console.log(diliveryPlanItems, "34567890")

        console.log(JSON.stringify(sortedData), "revLog")
        const revLog = formater(sortedData);

        console.log(revLog, "revLog refactored")
        console.log(respo, "----respo----")
        var uniqueArr = []
        var sec_arr = []
        respo.forEach((item) => {
          const exist_ = revLog.find((x) => x.LessonID === item.LessonID);
          if (exist_) {
            exist_["lesson"] = [...item.responses, ...exist_["lesson"]]
            console.log([...item.responses, ...exist_["lesson"]], "exist_")
            uniqueArr.push({
              ...exist_
            })
            sec_arr.push({
              ...exist_
            })
            console.log(exist_, 'wxist react')
          } else {
            uniqueArr.push({
              LessonID: item.LessonID,
              Project: item.Project,
              l_name: item.l_name,
              lesson: [...item.responses]
            })
            sec_arr.push({
              LessonID: item.LessonID,
              Project: item.Project,
              l_name: item.l_name,
              lesson: [...item.responses]
            })
          }
        })


        //write function to just rearrange item accordingly



        console.log(sec_arr, "before")
        console.log(uniqueArr, "before")
        const datauniqueArr = uniqueArr;
        datauniqueArr.map((res, index) => {

          const lessonData = datauniqueArr[index].lesson; //main array
          const responseFormator = revManager(lessonData)
          datauniqueArr[index].lesson = responseFormator
        })
        console.log(datauniqueArr, "after")
        setFinalStep(datauniqueArr)
        setFinalStepConst(datauniqueArr)
        setAdpLoader2("noLoader");
        console.log("no loader active")








      })
      .catch((err) => {
        console.log(err, "error in review log");
      });




  }
  // function filterLessons(data) {
  //   var stepsFilter = ["Draft", "Final Draft"];

  //   // Prepare an empty array to hold unique LessonIDs.
  //   var uniqueIds = [];

  //   // Loop over data, to add all unique LessonIDs into uniqueIds array
  //   for (var i = 0; i < data.length; i++) {
  //     if (uniqueIds.indexOf(data[i].LessonID) === -1) {
  //       uniqueIds.push(data[i].LessonID);
  //     }
  //   }

  //   var result = [];

  //   // Iterate over uniqueIds array
  //   for (var j = 0; j < uniqueIds.length; j++) {
  //     // Prepare an array for storing matching responses
  //     var responses = [];

  //     // Loop over the data again and add objects with matching LessonID and Steps
  //     for (var k = 0; k < data.length; k++) {
  //       if (data[k].LessonID === uniqueIds[j] && stepsFilter.indexOf(data[k].Steps) !== -1) {
  //         responses.push(data[k]);
  //       }
  //     }

  //     // Add responses array to the result array with LessonID.
  //     result.push({
  //       LessonID: uniqueIds[j],
  //       responses: responses
  //     });
  //   }

  //   return result;
  // }


  var objectRef = [ 
    {
      type: "Professional Learning (Lessons)",
      draft: "Write lesson outline complete"

    },
    {
      type: "Professional Learning (Practice Lessons)",
      draft: "Write lesson outline complete"
    },
    {
      type: "Event",
      draft: "Event brief Marketing brief Budget"

    }
    ,
    {
      type: "Professional Learning (Survey)",
      draft: "Draft"

    }, {
      type: "Curriculum (Survey)",
      draft: "Draft"
    },
    {
      type: "Marketing (Survey)",
      draft: "Draft"
    },
    {
      type: "Content Creation (Survey)",
      draft: "Draft"
    },
    {
      type: "Curriculum (Writing a lesson)",
      draft: "Draft"
    },
    {
      type: "Marketing (Starting a marketing campaign)",
      draft: "Produce a product board"
    },
    {
      type: "Marketing (Creating marketing collateral)",
      draft: "Draft copy"
    },
    {
      type: "Marketing (Delivering marketing collateral)",
      draft: "Review signed off campaign strategy"
    },
    {
      type: "Marketing (Promoting through the media)",
      draft: "Build media kit"
    },
    {
      type: "Marketing (Deliver logistics of events)",
      draft: "Event brief"
    },



    {
      type: "Content Creation (Sourcing digital content)",
      draft: "Final draft"
    },
    {
      type: "Content Creation (Build a video script)",
      draft: "Draft script"
    },
    {
      type: "Content Creation (Compile video from footage)",
      draft: "Additional drafts"
    },
    {
      type: "Content Creation (Shoot video footage)",
      draft: "Film new content"
    },
    {
      type: "Content Creation (Small graphic)",
      draft: "Draft"
    },
    {
      type: "Content Creation (Medium graphic)",
      draft: "Draft"
    },
    {
      type: "Content Creation (Digital Database)",
      draft: "Final draft"
    },
    {
      type: "Content Creation (Video Content Producer)",
      draft: "Draft"
    },



    {
      type: "Curriculum (Oz-e-Maths Swap Outs)",
      draft: "Draft"
    },
    {
      type: "Curriculum (General)",
      draft: "Draft"
    },
    {
      type: "School Improvement (SCM)",
      draft: "Draft development"
    },
    {
      type: "Curriculum (Overview)",
      draft: "Draft"
    },
    {
      type: "School Partnerships",
      draft: "Analyse data"
    },
    {
      type: "Curriculum Teaching Guide",
      draft: "Draft"
    },
    {
      type: "Curriculum Student Workbook",
      draft: "Draft"
    },
    {
      type: "Content Creation (Remote filming)",
      draft: "Review Video requirements form with Developer"
    },


    {
      type: "Professional Learning (Modules)",
      draft: "Additional drafts"

    },
    {
      type: "Business Services (Delivery a board meeting)",
      draft: "CEO to review draft minutes"

    },
    {
      type: "Content Creation (Large graphic)",
      draft: "Draft"

    },
    {
      type: "School Partnerships V2",
      draft: "Analyse data and write WDRR"

    }
  ]

  function getString(input: string | string[]): string {
    // Check if input is an array
    if (Array.isArray(input)) {
      // If input is an array, loop through the elements
      for (let i = 0; i < input.length; i++) {
        // If the element is a string, return it
        if (typeof input[i] === 'string') {
          return input[i];
        }
      }
    }

    // If input is a string, return it
    if (typeof input === 'string') {
      return input;
    }

    // If no string is found, return an empty string
    return '';
  }

  function getUniqueLessonsWithSteps(data) {
    console.log(data, "doiilvsdf")
    let unique_arr = []
    data.forEach(element => {
      if (unique_arr.findIndex(x => x.LessonID === element.LessonID) === -1) {
        unique_arr.push({
          LessonID: element.LessonID,

          Project: element.Project,
          l_name: element.Lesson,
          responses: []
        })
      }
    });

    data.forEach(x => {
      console.log("----------------------------------")
      console.log(x)
      const TypeStringConverted = getString(x.Types)
      console.log(TypeStringConverted)

      const getFruit = objectRef.find(fruit => fruit.type === TypeStringConverted);
      console.log(getFruit)

      if (getFruit && getFruit.draft !== "no" && getFruit.type === TypeStringConverted && getFruit.draft === x.Steps) {

        unique_arr[unique_arr.findIndex(y => y.LessonID === x.LessonID)].responses.push({


          Dev: x?.Developer?.name,
          Dev2: x?.Developer,
          End: x.End, FromEmail: x?.Developer?.email,
          Lesson: x.Lesson, LessonID: x.LessonID, Project: x.Project, Start: x.Start,
          ToEmail: x?.Developer?.email,
          ID: 0,
          ADPId: x.ID,
          request: "Draft",
          response: x.Steps


        })

      }


    })

    console.log(unique_arr, "unique_arr")


    return unique_arr
  }




  const hoursCal = (data) => {
    data.forEach((lessonData) => {
      lessonData.forEach((item) => {

        const hours = calculateWeekdaysWithHours(new Date(
          moment(item?.Start, DateListFormat).format(DatePickerFormat)
        ), new Date(
          moment(item?.End, DateListFormat).format(DatePickerFormat)
        ));
        setTotalHours(prevTotalHours => prevTotalHours + hours);
      });
    });
  }
  window.onbeforeunload = function (e) {
    if (adpAutoSave) {
      let dialogText =
        "You have unsaved changes, are you sure you want to leave?";
      e.returnValue = dialogText;
      return dialogText;
    }
  };

  // States-Declaration Ends
  //Function-Section Starts
  const generateExcel = () => {



    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('Horizontal Table');
    // Add Steps row

    sheet.columns = [
      { header: "Task", key: "task", width: 25 },
      { header: "Activity", key: "activity", width: 25 },

    ];

    StepsArray.forEach(step => {
      const cell = sheet.getCell(1, sheet.columnCount + 1);
      cell.value = step;
      sheet.mergeCells(1, cell.col, 1, cell.col + 1); // Apply colspan to merged cells
    });
    // sheet.addRow(stepsRow);
    Tabledata.map((data, index) => {


      const headerRow = [index + 1 == Tabledata.length ? data[0].Types : '', data[0].Lesson];
      data.forEach(nested => {
        headerRow.push(nested?.Start, nested?.End);
      });
      sheet.addRow(headerRow);
      const PhotoRow = ['', ''];
      data.forEach(nested => {
        PhotoRow.push(nested?.Developer?.name, '');
      });
      sheet.addRow(PhotoRow);

    })
    workbook.xlsx.writeBuffer().then(buffer => {
      const file = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      FileSaver.saveAs(file, 'horizontal_table.xlsx');
    });
  };
  const calculateWeekdays = (startDate, endDate) => {


    startDate = convertISODateToStartDate(startDate)
    endDate = convertISODateToStartDate(endDate)


    // Validate input
    if (endDate < startDate)
      return 0;

    // Calculate days between dates
    var millisecondsPerDay = 86400 * 1000; // Day in milliseconds
    startDate.setHours(0, 0, 0, 1);  // Start just after midnight
    endDate.setHours(23, 59, 59, 999);  // End just before midnight
    var diff = endDate - startDate;  // Milliseconds between datetime objects
    var days = Math.ceil(diff / millisecondsPerDay);

    // Subtract two weekend days for every week in between
    var weeks = Math.floor(days / 7);
    days = days - (weeks * 2);

    // Handle special cases
    var startDay = startDate.getDay();
    var endDay = endDate.getDay();

    // Remove weekend not previously removed.
    if (startDay - endDay > 1)
      days = days - 2;

    // Remove start day if span starts on Sunday but ends before Saturday
    if (startDay == 0 && endDay != 6) {
      days = days - 1;
    }

    // Remove end day if span ends on Saturday but starts after Sunday
    if (endDay == 6 && startDay != 0) {
      days = days - 1;
    }

    return days;
  }

  const convertISODateToStartDate = (isoDate) => {
    // Create a JavaScript Date object from the provided ISO 8601 formatted string
    let date = new Date(isoDate);

    // Format the date to the required "DatePickerFormat"
    let formattedDate = moment(date).format("YYYY-MM-DDT14:00:00Z");

    // Convert back to JavaScript Date object and return
    return new Date(formattedDate);
  };
  //   const calculateWeekdaysWithHours = (startDate, endDate) => {
  // // Validate input
  // if (endDate < startDate)
  // return 0;

  // // Calculate hours between dates
  // var millisecondsPerHour = 60 * 60 * 1000; // Hour in milliseconds

  // var diff = endDate - startDate;  // Milliseconds between datetime objects
  // var totalHours = Math.ceil(diff / millisecondsPerHour);

  // // Calculate start and end day of the week
  // var startDay = startDate.getDay();
  // var endDay = endDate.getDay();

  // // Adjust total hours based on weekends
  // if (startDay === 0)
  // startDay = 7; // Sunday is considered as day 7

  // if (endDay === 0)
  // endDay = 7; // Sunday is considered as day 7

  // var weekends = Math.floor((totalHours + startDay - 1) / 24 / 7) * 2;

  // // Adjust start and end hours
  // var startHour = startDate.getHours();
  // var endHour = endDate.getHours();

  // if (startDay !== 6 && startDay !== 7) {
  // if (startHour > 0 && startHour < 24) {
  //   totalHours -= startHour;
  //   weekends--;
  // }
  // }

  // if (endDay !== 6 && endDay !== 7) {
  // if (endHour > 0 && endHour < 24) {
  //   totalHours -= 24 - endHour;
  //   weekends--;
  // }
  // }

  // // Subtract weekends hours from the total hours
  // totalHours -= weekends * 24;



  // return totalHours +"hrs";
  //   }
  const calculateWeekdaysWithHours = (startDate, endDate) => {
    // Validate input
    if (endDate < startDate)
      return 0;

    // Calculate hours and days between dates
    var millisecondsPerHour = 60 * 60 * 1000; // Hour in milliseconds
    var millisecondsPerDay = 24 * millisecondsPerHour; // Day in milliseconds

    startDate.setHours(0, 0, 0, 1);  // Start just after midnight
    endDate.setHours(23, 59, 59, 999);  // End just before midnight

    var diff = endDate - startDate;  // Milliseconds between datetime objects
    var totalHours = Math.ceil(diff / millisecondsPerHour);
    var totalDays = Math.ceil(diff / millisecondsPerDay);

    // Subtract two weekend days for every week in between
    var weeks = Math.floor(totalDays / 7);
    totalDays = totalDays - (weeks * 2);

    // Handle special cases
    var startDay = startDate.getDay();
    var endDay = endDate.getDay();

    // Remove weekends not previously removed
    if (startDay - endDay > 1)
      totalDays = totalDays - 2;

    // Remove start day if span starts on Sunday but ends before Saturday
    if (startDay === 0 && endDay !== 6) {
      totalDays = totalDays - 1;
      totalHours = totalHours - (24 - startDate.getHours());
    }

    // Remove end day if span ends on Saturday but starts after Sunday
    if (endDay === 6 && startDay !== 0) {
      totalDays = totalDays - 1;
      totalHours = totalHours - endDate.getHours() + 1;
    }

    // Adjust hours for complete days
    totalHours = totalHours - (totalDays * 24);
    let time = totalDays * 24 + totalHours;

    return time



  }



  const fetchAnnualPlanProduct = async (Project) => {
    let resp = {
      ba: '',
      term: 0
    }
    await sharepointWeb.lists.getByTitle("Annual Plan").items.filter(`Title eq '${Project}'`).get().then((items) => {
      if (items.length > 0) {
        // Item found
        const item = items[0];
        resp = {
          ba: item.BA_x0020_acronyms,
          term: item.TermNew.join(', ')
        }

        setAnnualPlanData(resp)
        // Your further code logic here...
      } else {
        // Item not found 
      }
    }).catch((error) => {
      console.log("Error occurred:", error);
    });

    return resp;
  }


  const adpGetCurrentUserDetails = () => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let adpCurrentUser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser({ ...adpCurrentUser });
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetCurrentUserDetails");
      });
  };
  const getActivityPlanItem = async () => {
    let _adpItem = [];

    sharepointWeb.lists
      .getByTitle(activityPlanListName)
      .items.getById(activityPlan_ID)
      .get()
      .then((item) => {


        //create function to fetch annual plan product by title fro annul and project from activ plan




        fetchAnnualPlanProduct(item.Project)



        _adpItem.push({
          ID: item.Id ? item.Id : "",
          Lesson: item.Lessons ? item.Lessons : "",
          Project: item.Project ? item.Project : "",
          Product: item.Product ? item.Product : "",
          ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
          ProductVersion: item.ProductVersion ? item.ProductVersion : "V1",
          Types: item.Types ? item.Types : "",
          Title: item.Title ? item.Title : "",


          Status: item.Status ? item.Status : null,
        });

        let _adpLessons = [];
        let lessons = _adpItem[0].Lesson.split(";");

        lessons.forEach((ls) => {
          _adpLessons.push({
            ID: ls.split("~")[0],
            Name: ls.split("~")[1],
            StartDate: ls.split("~")[2],
            EndDate: ls.split("~")[3],
            DeveloperId:
              ls.split("~")[4] != "NaN" && ls.split("~")[4] != null
                ? parseInt(ls.split("~")[4])
                : null,
            DeveloperName:
              ls.split("~")[4] != "NaN" &&
                ls.split("~")[4] != "null" &&
                allPeoples.length > 0 &&
                allPeoples.filter((ap) => {
                  return ap.ID == ls.split("~")[4];
                }).length > 0
                ? allPeoples.filter((ap) => {
                  return ap.ID == ls.split("~")[4];
                })[0].text
                : null,
            DeveloperEmail:
              ls.split("~")[4] != "NaN" &&
                ls.split("~")[4] != "null" &&
                allPeoples.length > 0 &&
                allPeoples.filter((ap) => {
                  return ap.ID == ls.split("~")[4];
                }).length > 0
                ? allPeoples.filter((ap) => {
                  return ap.ID == ls.split("~")[4];
                })[0].secondaryText
                : null,
          });
        });







        adpGetData(_adpItem[0], _adpLessons);
        setActivtyPlanItem([..._adpItem]);
      })
      .catch((err) => {
        adpErrorFunction(err, "getActivityPlanItem");
      });
  };
  const getActivityPBData = () => {
    sharepointWeb.lists
      .getByTitle(activityPBListName)
      .items.filter(
        `ActivityPlanID eq '${activityPlan_ID}'
        and Week eq '${adpCurrentWeekNumber}'
        and Year eq '${adpCurrentYear}'`
      )
      .top(5000)
      .get()
      .then((items) => {
        setActivityPB([...items]);
      })
      .catch((err) => {
        adpErrorFunction(err, "getActivityPBData");
      });
  };
  const adpGetData = (adpItem: any, lessons) => {
    let adpAllitems = [];
    sharepointWeb.lists
      .getByTitle(adpListName)
      .items.select(
        "*",
        "Developer/Title",
        "Developer/Id",
        "Developer/EMail",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/EndDate"
      )
      .expand("Developer,FieldValuesAsText")
      .filter(`ActivityPlanID eq ${activityPlan_ID}`)
      .orderBy("OrderId", true)
      .top(5000)
      .get()
      .then((items) => {
        if (items.length > 0) {
          console.log(items, "activity planner")
          items.forEach((item, index) => {
            adpAllitems.push({
              OrderId: index,
              LessonID: item.LessonID,
              ID: item.Id ? item.Id : "",
              Steps: item.Title ? item.Title : "",
              PH: item.PlannedHours ? item.PlannedHours : "",
              MinPH: item.MinPH ? item.MinPH : "",
              EditorId: item.EditorId ? item?.EditorId : "",
              MaxPH: item.MaxPH ? item.MaxPH : "",
              Project: item.Project ? item.Project : "",
              Lesson: item.Lesson ? item.Lesson : "",
              Types: item.Types ? item.Types : "",
              Title: item.Title ? item.Title : "",

              Start: item.StartDate
                ? moment(
                  item["FieldValuesAsText"].StartDate,
                  DateListFormat
                ).format(DateListFormat)
                : null,
              End: item.EndDate
                ? moment(
                  item["FieldValuesAsText"].EndDate,
                  DateListFormat
                ).format(DateListFormat)
                : null,
              Developer: item.DeveloperId
                ? {
                  name: item.Developer.Title,
                  id: item.Developer.Id,
                  email: item.Developer.EMail,
                }
                : "",
              Status: item.Status ? item.Status : "noData",
              IsCompleteStatus: item.Status == "Completed" ? true : false,
              IsCompleteNew: false,
              AH: item.ActualHours ? item.ActualHours : 0,
              dateError: false,
              PHError: false,
              PHWeek: item.PHWeek ? item.PHWeek : null,
            });
          });
          adpGetTemplateData(adpItem, lessons, adpAllitems, items.length);
          // groups(adpAllitems);
          // adpGetAllOptions(adpAllitems);

          // setAdpActivityResponseData([...adpAllitems]);
          // setAdpData([...adpAllitems]);
          // setAdpMasterData([...adpAllitems]);
          // setAdpLoader("noLoader");
        } else {
          setNewDataFlag(true);
          sharepointWeb.lists
            .getByTitle(templateListName)
            .items.filter(`Types eq '${adpItem.Types}'`)
            .orderBy("ID", true)
            .top(5000)
            .get()
            .then((items) => {
              let count = 0;
              lessons.forEach((ls) => {
                items.forEach((item, index) => {
                  let PHErrorFlag =
                    item.MinHours && item.MaxHours
                      ? adpPHValidationFunction(
                        parseFloat(item.Hours ? item.Hours : 0),
                        item.MinHours,
                        item.MaxHours
                      )
                      : false;
                  // let datediff =
                  //   new Date(ls.EndDate).getDate() -
                  //   new Date(ls.StartDate).getDate() +
                  //   1;
                  // let Hours = item.Week ? datediff / 7 : item.Hours;
                  const date1: any = new Date(ls.StartDate);
                  const date2: any = new Date(ls.EndDate);
                  const diffTime = Math.abs(date2 - date1);
                  const diffDays =
                    Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
                  const Hours = item.Week ? diffDays / 7 : item.Hours;

                  adpAllitems.push({
                    OrderId: count++,
                    ID: count,
                    Steps: item.Title ? item.Title : "",
                    Types: item.Types ? item.Types : "",
                    Title: item.Title ? item.Title : "",

                    PH: Hours ? Hours : "",
                    MinPH: item.MinHours ? item.MinHours : "",
                    MaxPH: item.MaxHours ? item.MaxHours : "",
                    Project: adpItem.Project ? adpItem.Project : "",
                    LessonID: ls.ID,
                    Lesson: ls.Name ? ls.Name : "",
                    Start: moment(ls.StartDate).format(DateListFormat),
                    End: moment(ls.EndDate).format(DateListFormat),
                    Developer: {
                      name: ls.DeveloperName,
                      id: ls.DeveloperId,
                      email: ls.DeveloperEmail,
                    },
                    Status: "Scheduled",
                    IsCompleteStatus: false,
                    IsCompleteNew: false,
                    AH: 0,
                    dateError: false,
                    PHError: PHErrorFlag,
                    PHWeek: item.Week ? item.Week : null,
                  });
                });
              });
              groups(adpAllitems);
              adpGetAllOptions(adpAllitems);

              // setAdpActivityResponseData([...adpAllitems]);
              setAdpData([...adpAllitems]);
              setAdpMasterData([...adpAllitems]);
              setAdpLoader("noLoader");
            })
            .catch((err) => {
              adpErrorFunction(err, "adpGetData-getTemplateData");
            });
        }
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetData-getADPData");
      });
  };

  //!Update template in the database
  const adpGetTemplateData = (
    adpItem: any,
    lessons,
    adplistItems: any[],
    countLists
  ) => {
    let adpAllitems = adplistItems;
    sharepointWeb.lists
      .getByTitle(templateListName)
      .items.filter(`Types eq '${adpItem.Types}'`)
      .orderBy("ID", true)
      .top(5000)
      .get()
      .then((items) => {
        let count = countLists;
        lessons.forEach((ls) => {
          let curLessondata = adplistItems.filter((arr) => {
            return arr.LessonID == ls.ID;
          });

          curLessondata.length == 0 &&
            items.forEach((item, index) => {
              let PHErrorFlag =
                item.MinHours && item.MaxHours
                  ? adpPHValidationFunction(
                    parseFloat(item.Hours ? item.Hours : 0),
                    item.MinHours,
                    item.MaxHours
                  )
                  : false;
              const date1: any = new Date(ls.StartDate);
              const date2: any = new Date(ls.EndDate);
              const diffTime = Math.abs(date2 - date1);
              const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
              const Hours = item.Week ? diffDays / 7 : item.Hours;

              adpAllitems.push({
                OrderId: count++,
                ID: 0,
                Steps: item.Title ? item.Title : "",
                PH: Hours ? Hours : "",
                MinPH: item.MinHours ? item.MinHours : "",
                MaxPH: item.MaxHours ? item.MaxHours : "",
                Project: adpItem.Project ? adpItem.Project : "",
                LessonID: ls.ID,
                Lesson: ls.Name ? ls.Name : "",
                Types: item.Types ? item.Types : "",
                Title: item.Title ? item.Title : "",

                Start: moment(ls.StartDate).format(DateListFormat),
                End: moment(ls.EndDate).format(DateListFormat),
                Developer: {
                  name: ls.DeveloperName,
                  id: ls.DeveloperId,
                  email: ls.DeveloperEmail,
                },
                Status: "Scheduled",
                IsCompleteStatus: false,
                IsCompleteNew: false,
                AH: 0,
                dateError: false,
                PHError: PHErrorFlag,
                PHWeek: item.Week ? item.Week : null,
              });
            });
        });
        groups(adpAllitems);
        adpGetAllOptions(adpAllitems);

        // setAdpActivityResponseData([...adpAllitems]);
        setAdpData([...adpAllitems]);
        setAdpMasterData([...adpAllitems]);
        setAdpLoader("noLoader");
      })
      .catch((err) => {
        adpErrorFunction(err, "adpGetData-getTemplateData");
      });
  };

  const adpGetAllOptions = (allItems: any) => {
    allItems.forEach((item: any) => {
      if (
        adpDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
          return developerOptn.key == item.Developer.name;
        }) == -1 &&
        item.Developer.name
      ) {
        adpDrpDwnOptns.developerOptns.push({
          key: item.Developer.name,
          text: item.Developer.name,
        });
      }

      if (
        adpDrpDwnOptns.stepsOptns.findIndex((stepsOptn) => {
          return stepsOptn.key == item.Steps;
        }) == -1 &&
        item.Steps
      ) {
        adpDrpDwnOptns.stepsOptns.push({
          key: item.Steps,
          text: item.Steps,
        });
      }
      if (
        adpDrpDwnOptns.statusOptns.findIndex((statsOptn) => {
          return statsOptn.key == item.Status;
        }) == -1 &&
        item.Status
      ) {
        adpDrpDwnOptns.statusOptns.push({
          key: item.Status,
          text: item.Status,
        });
      }

      if (
        adpDrpDwnOptns.lessonOptns.findIndex((lessonOptn) => {
          return lessonOptn.key == item.Lesson;
        }) == -1 &&
        item.Lesson
      ) {
        adpDrpDwnOptns.lessonOptns.push({
          key: item.Lesson,
          text: item.Lesson,
        });
      }
    });

    let maxWeek =
      parseInt(adpFilters.year) == moment().year() ? moment().isoWeek() : 53;

    for (var i = 1; i <= maxWeek; i++) {
      adpDrpDwnOptns.weekOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }
    for (var i = 2020; i <= moment().year(); i++) {
      adpDrpDwnOptns.yearOptns.push({
        key: i.toString(),
        text: i.toString(),
      });
    }

    let unsortedFilterKeys = adpSortingFilterKeys(adpDrpDwnOptns);
    setAdpDropDownOptions({ ...unsortedFilterKeys });
  };
  const adpSortingFilterKeys = (unsortedFilterKeys: any) => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    if (
      unsortedFilterKeys.developerOptns.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      unsortedFilterKeys.developerOptns.shift();
      let loginUserIndex = unsortedFilterKeys.developerOptns.findIndex(
        (user) => {
          return (
            user.text.toLowerCase() ==
            props.spcontext.pageContext.user.displayName.toLowerCase()
          );
        }
      );
      let loginUserData = unsortedFilterKeys.developerOptns.splice(
        loginUserIndex,
        1
      );

      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift(loginUserData[0]);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    } else {
      unsortedFilterKeys.developerOptns.shift();
      unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
      unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
    }

    unsortedFilterKeys.statusOptns.shift();
    unsortedFilterKeys.statusOptns.sort(sortFilterKeys);
    unsortedFilterKeys.statusOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.stepsOptns.shift();
    unsortedFilterKeys.stepsOptns.sort(sortFilterKeys);
    unsortedFilterKeys.stepsOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.lessonOptns.shift();
    unsortedFilterKeys.lessonOptns.sort(sortFilterKeys);
    unsortedFilterKeys.lessonOptns.unshift({ key: "All", text: "All" });

    return unsortedFilterKeys;
  };
  const adpListFilter = (key: string, option: any) => {
    let arrBeforeFilter = [...adpData];

    let tempFilterKeys = { ...adpFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.developer != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Developer.name == tempFilterKeys.developer;
      });
    }

    if (tempFilterKeys.step != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Steps == tempFilterKeys.step;
      });
    }
    if (tempFilterKeys.status != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Status == tempFilterKeys.status;
      });
    }
    if (tempFilterKeys.lesson) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Lesson.toLowerCase().includes(
          tempFilterKeys.lesson.toLowerCase()
        );
      });
    }

    if (tempFilterKeys.week != "All") {
      let year =
        tempFilterKeys.year == "All" ? moment().year() : tempFilterKeys.year;

      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        let start = moment(arr.Start, DateListFormat)
          .year()
          .toString()
          .concat(
            (
              "0" + moment(arr.Start, DateListFormat).isoWeek().toString()
            ).slice(-2)
          );
        let end = moment(arr.End, DateListFormat)
          .year()
          .toString()
          .concat(
            ("0" + moment(arr.End, DateListFormat).isoWeek().toString()).slice(
              -2
            )
          );
        let today = year
          .toString()
          .concat(("0" + tempFilterKeys.week.toString()).slice(-2));

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }

    if (tempFilterKeys.year != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        let start = moment(arr.Start, DateListFormat).year().toString();

        let end = moment(arr.End, DateListFormat).year().toString();

        let today = tempFilterKeys.year.toString();

        return (
          parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
        );
      });
    }

    groups([...arrBeforeFilter]);
    // setAdpActivityResponseData([...arrBeforeFilter]);
    setAdpFilters({ ...tempFilterKeys });
  };
  const overallPlannedHours = () => {
    let ph = 0;
    if (adpData.length > 0) {
      adpData.forEach((data) => {
        ph += data.PH ? data.PH : 0;
      });
    }
    return ph;
  };
  const overallActualHours = () => {
    let ah = 0;
    if (adpData.length > 0) {
      adpData.forEach((data) => {
        ah += data.AH ? data.AH : 0;
      });
    }
    return ah;
  };

  const updateDeveloperData = () => {

  }


  const adpActivityResponseHandler = (id: number, key: string, value: any) => {
    //////.......................................... 


    let tempDeveloper = [];
    console.log(adpData, "adpActivityResponse")
    let Index = adpData.findIndex((data) => data.OrderId == id);
    let disIndex = adpActivityResponseData.findIndex(
      (data) => data.OrderId == id
    );

    let adpBeforeData = adpData[Index];

    if (key == "Developer") {
      if (value) {
        tempDeveloper = allPeoples.filter((people) => {
          return people.ID == value;
        });
      }
    }

    let dateErrorFlag = adpDateValidationFunction(
      key == "Start"
        ? moment(value).format("YYYY/MM/DD")
        : moment(adpBeforeData.Start, DateListFormat).format("YYYY/MM/DD"),
      key == "End"
        ? moment(value).format("YYYY/MM/DD")
        : moment(adpBeforeData.End, DateListFormat).format("YYYY/MM/DD")
    );
    let PHErrorFlag =
      key == "PH"
        ? adpPHValidationFunction(
          parseFloat(value),
          adpBeforeData.MinPH,
          adpBeforeData.MaxPH
        )
        : adpBeforeData.PHError;

    const date1: any = new Date(key == "Start" ? value : adpBeforeData.Start);
    const date2: any = new Date(key == "End" ? value : adpBeforeData.End);
    const diffTime = Math.abs(date2 - date1);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    const Hours = adpBeforeData.PHWeek ? diffDays / 7 : adpBeforeData.PH;

    let adpOnchangeData = {
      LessonID: adpBeforeData.LessonID,
      OrderId: adpBeforeData.OrderId,
      ID: adpBeforeData.ID,
      Steps: adpBeforeData.Steps,
      PH: key == "PH" ? value : Hours,
      MinPH: adpBeforeData.MinPH,
      MaxPH: adpBeforeData.MaxPH,
      Project: adpBeforeData.Project,
      Lesson: adpBeforeData.Lesson,
      Start:
        key == "Start"
          ? moment(value).format(DateListFormat)
          : adpBeforeData.Start,
      End:
        key == "End" ? moment(value).format(DateListFormat) : adpBeforeData.End,
      Developer:
        key == "Developer"
          ? {
            name: tempDeveloper.length > 0 ? tempDeveloper[0].text : "",
            id: tempDeveloper.length > 0 ? tempDeveloper[0].ID : null,
            email:
              tempDeveloper.length > 0 ? tempDeveloper[0].secondaryText : "",
          }
          : adpBeforeData.Developer,
      Status: adpBeforeData.Status,
      IsCompleteStatus:
        key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteStatus,
      IsCompleteNew:
        key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteNew,
      dateError: dateErrorFlag,
      PHError: PHErrorFlag,
      PHWeek: adpBeforeData.PHWeek,
    };

    adpData[Index] = adpOnchangeData;
    adpActivityResponseData[disIndex] = adpOnchangeData;

    let isCompleteDetails = adpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });
    adpData.length == isCompleteDetails.length
      ? setAdpIsCompleted(true)
      : setAdpIsCompleted(false);

    setAdpData([...adpData]);
    groups([...adpActivityResponseData]);
    // setAdpActivityResponseData([...adpActivityResponseData]);
  };
  const adpAddItem = () => {
    let successCount = 0;

    let completionDetails = adpData.filter((arr) => {
      return arr.IsCompleteStatus == true;
    });

    let completionValue =
      adpData.length > 0 && completionDetails.length > 0
        ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
        : 0;

    adpData.forEach(async (response: any, index: number) => {
      let strDSNA: string = `${response.Developer.id ? response.Developer.id : null
        }-0`;

      let statusValue = response.IsCompleteStatus
        ? "Completed"
        : response.Status
          ? response.Status
          : null;

      let responseData = {
        ActivityPlanID: activityPlan_ID ? activityPlan_ID.toString() : "",
        Title: response.Steps ? response.Steps : "",
        PlannedHours: response.PH ? response.PH : 0,
        MinPH: response.MinPH ? response.MinPH : 0,
        MaxPH: response.MaxPH ? response.MaxPH : 0,
        ProjectVersion: activtyPlanItem[0].ProjectVersion
          ? activtyPlanItem[0].ProjectVersion
          : "V1",
        ProductVersion: activtyPlanItem[0].ProductVersion
          ? activtyPlanItem[0].ProductVersion
          : "V1",
        Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
        Types: activtyPlanItem[0].Types ? activtyPlanItem[0].Types : "",



        Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
        Lesson: response.Lesson ? response.Lesson : "",
        StartDate: response.Start
          ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
          : moment().format("YYYY-MM-DD"),
        EndDate: response.End
          ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
          : moment().format("YYYY-MM-DD"),
        DeveloperId: response.Developer.id ? response.Developer.id : null,
        // Status: "Scheduled",
        Status: statusValue,
        ActualHours: 0,
        OrderId: response.OrderId,
        LessonID: response.LessonID ? response.LessonID : null,
        PHWeek: response.PHWeek ? response.PHWeek : null,
        SPFxFilter: strDSNA,
      };

      // debugger;

      await sharepointWeb.lists
        .getByTitle(adpListName)
        .items.add(responseData)
        .then((item) => {
          successCount++;
          adpData[index].ID = item.data.Id;
          adpData[index].Status = statusValue;
          adpData[index].IsCompleteNew = false;

          if (adpData.length == successCount) {
            let apCompletedValue = adpData.filter((arr) => {
              return arr.IsCompleteStatus == true;
            });
            if (
              adpData.length == apCompletedValue.length &&
              activtyPlanItem[0].Status != "Completed"
            ) {
              sharepointWeb.lists
                .getByTitle(activityPlanListName)
                .items.getById(activityPlan_ID)
                .update({
                  Status: "Completed",
                  Completion: 100,
                  CompletedDate: moment().format("YYYY-MM-DD"),
                })
                .then((e) => { })
                .catch((err) => {
                  adpErrorFunction(err, "saveDPData-getAPItem");
                });
            } else {
              sharepointWeb.lists
                .getByTitle(activityPlanListName)
                .items.getById(activityPlan_ID)
                .update({
                  Completion: completionValue,
                })
                .then((e) => { })
                .catch((err) => {
                  adpErrorFunction(err, "saveDPData-getAPItem");
                });
            }

            const newData = _copyAndSort(adpData, "OrderId", false);

            adpGetAllOptions([...newData]);
            setAdpMasterData([...newData]);
            setAdpData([...newData]);
            groups([...newData]);

            // setAdpActivityResponseData([...adpData]);
            setNewDataFlag(false);
            setAdpItemAddFlag(true);
            setAdpEditFlag(false);
            setAdpLoader("noLoader");
            AddSuccessPopup();
          }
        })
        .catch((err) => {
          adpErrorFunction(err, "adpAddItem");
        });
    });
  };

  const adpUpdateItem_Old = () => {
    let responseDataArr = [];
    let newArr = [...adpData];
    let successCount = 0;

    let selected = [];

    adpData.forEach((response: any, index: number) => {
      let targetStatus = newArr.filter((arr) => {
        return arr.ID == response.ID;
      });

      let strDSNA: string = `${response.Developer.id}-${targetStatus[0].Status == "Completed" ? 1 : 0
        }`;

      let responseData = {
        ProjectVersion: activtyPlanItem[0].ProjectVersion
          ? activtyPlanItem[0].ProjectVersion
          : "V1",
        ProductVersion: activtyPlanItem[0].ProductVersion
          ? activtyPlanItem[0].ProductVersion
          : "V1",
        Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
        Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
        PlannedHours: response.PH ? response.PH : 0,
        StartDate: response.Start
          ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
          : null,
        EndDate: response.End
          ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
          : null,
        DeveloperId: response.Developer.id ? response.Developer.id : null,
        SPFxFilter: strDSNA,
      };

      responseDataArr.push(responseData);

      sharepointWeb.lists
        .getByTitle(adpListName)
        .items.getById(response.ID)
        .update(responseData)
        .then(() => {
          successCount++;
          let newDeveloperDetails = {};

          let targetIndex = newArr.findIndex((arr) => arr.ID == response.ID);
          let targetItem = newArr.filter((arr) => {
            return arr.ID == response.ID;
          });

          if (response.Developer.id) {
            let newDeveloper = allPeoples.filter((people) => {
              return people.ID == response.Developer.id;
            });
            newDeveloperDetails = {
              name: newDeveloper[0].text,
              id: newDeveloper[0].ID,
              email: newDeveloper[0].secondaryText,
            };
          } else {
            newDeveloperDetails = {
              name: null,
              id: null,
              email: null,
            };
          }

          newArr[targetIndex] = {
            OrderId: response.OrderId,
            ID: targetItem[0].ID ? targetItem[0].ID : "",
            Steps: targetItem[0].Steps ? targetItem[0].Steps : "",
            PH: response.PH ? response.PH : "",
            MinPH: targetItem[0].MinPH ? targetItem[0].MinPH : "",
            MaxPH: targetItem[0].MaxPH ? targetItem[0].MaxPH : "",
            Project: targetItem[0].Project ? targetItem[0].Project : "",
            LessonID: targetItem[0].LessonID ? targetItem[0].LessonID : null,
            Lesson: targetItem[0].Lesson ? targetItem[0].Lesson : "",
            Start: response.Start ? response.Start : targetItem[0].Start,
            End: response.End ? response.End : targetItem[0].End,
            Developer: newDeveloperDetails,
            Status: targetItem[0].Status ? targetItem[0].Status : "",
            AH: targetItem[0].AH ? targetItem[0].AH : "",
            dateError: false,
            PHError: false,
            PHWeek: targetItem[0].PHWeek ? targetItem[0].PHWeek : null,
          };

          let filteredPB = activityPB.filter((pb) => {
            return pb.ActivityDeliveryPlanID == newArr[targetIndex].ID;
          });

          selected.push([...filteredPB]);

          if (filteredPB.length > 0) {
            sharepointWeb.lists
              .getByTitle(activityPBListName)
              .items.getById(filteredPB[0].ID)
              .update({
                PlannedHours: response.PH ? response.PH : 0,
                StartDate: response.Start
                  ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
                  : null,
                EndDate: response.End
                  ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
                  : null,
                DeveloperId: response.Developer.id
                  ? response.Developer.id
                  : null,
              })
              .then((e) => { })
              .catch((err) => {
                adpErrorFunction(err, "adpUpdateItem-updateAPBList");
              });
          }

          if (adpActivityResponseData.length == successCount) {
            adpGetAllOptions(newArr);
            setAdpEditFlag(false);
            setAdpMasterData([...newArr]);
            setAdpLoader("noLoader");
            AddSuccessPopup();
          }
        })
        .catch((err) => {
          adpErrorFunction(err, "adpUpdateItem-updateATPList");
        });
    });
  };
  const show = (position) => {
    let data = Open.filter((i) => i === position);

    if (data.length > 0) {
      setOpen(Open.filter((i) => i != position));
    } else {
      setOpen([...Open, position]);
    }

  };

  function convertDateFormat(inputDate) {
    const formattedDate = moment(inputDate).utc().format('YYYY-MM-DDTHH:mm:ss[Z]');
    return formattedDate;
  }

  function convertToCustomFormat(inputDate, outputFormat) {
    const formattedDate = moment(inputDate).format(outputFormat);
    return formattedDate;
  }
  const adpUpdateItem = () => {
    // let responseDataArr = [];
    // let newArr = [...adpData];
    // let successCount = 0;

    // let completionDetails = adpData.filter((arr) => {
    //   return arr.IsCompleteStatus == true;
    // });

    // let completionValue =
    //   adpData.length > 0 && completionDetails.length > 0
    //     ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
    //     : 0;
    // console.log(diliveryPlanNeedToBeUpdate)
    console.log(diliveryPlanNeedToBeUpdate), "1234567890"
    diliveryPlanNeedToBeUpdate.forEach((item, index) => {
      if (item.ADPId) {

        let inputDate1
        if (item.Start) {
          inputDate1 = item.Start
        } else {
          inputDate1 = ''
        }

        const inputFormat1 = "DD/MM/YYYY";
        const outputFormat1 = "YYYY-MM-DDTHH:mm:ss[Z]";

        const convertedDate1 = convertDateFormat2(inputDate1, inputFormat1, outputFormat1);

        let inputDate2;
        if (item.End) {
          inputDate2 = item.End
        } else {
          inputDate2 = ''
        }




        const inputFormat2 = "DD/MM/YYYY";
        const outputFormat2 = "YYYY-MM-DDTHH:mm:ss[Z]";

        const convertedDate2 = convertDateFormat2(inputDate2, inputFormat2, outputFormat2);




        let responseData2 = {
          // ProjectVersion: "v2",
          // ProductVersion: "v2",
          // Product: "Mastery Teaching Pathway",
          // Project: "Coach Positive High Expectations ",

          // PlannedHours: 0,
          StartDate: convertedDate1,
          EndDate: convertedDate2,
          DeveloperId: item?.Dev2?.id,


        }


        sharepointWeb.lists
          .getByTitle(adpListName)
          .items.getById(item.ADPId)
          .update(responseData2)
          .then(() => {
            console.log("success");


          })
          .catch((err) => {
            adpErrorFunction(err, "adpUpdateItem-updateATPList");
          });
      }
    })

    setFinalStepConst(finalStep)
    setAdpEditFlag(false);
    setAdpAutoSave(false);

  };









  const adpDateValidationFunction = (startDate: any, EndDate: any) => {
    if (startDate != null && EndDate != null) {
      if (startDate > EndDate) {
        return true;
      } else {
        return false;
      }
    } else {
      return false;
    }
  };
  const adpPHValidationFunction = (val, min, max) => {
    if (val >= min && val <= max) {
      return false;
    } else {
      return true;
    }
  };


  function convertDateFormat2(inputDate, inputFormat, outputFormat) {
    const parsedDate = moment(inputDate, inputFormat);
    const formattedDate = parsedDate.utc().format(outputFormat);
    return formattedDate;
  }



  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };
  const GetUserDetails = (filterText, currentPersonas) => {
    let _allPeoples = allPeoples;

    _allPeoples = _allPeoples.filter((arr) => {
      return arr.text.toLowerCase().indexOf("archive") == -1;
    });

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (arr) => !currentPersonas.some((persona) => persona.ID == arr.ID)
      );
    }

    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };
  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const adpErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Activity delivery plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setAdpLoader("noLoader");
        setAdpEditFlag(false);
        ErrorPopup();
        setAdpReRender(!adpReRender);
      }
    );
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity planner is successfully submitted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  const sortingFunction = (columnName, sortType): void => {
    let tempArr = adpData;
    let tempDisArr = adpActivityResponseData;

    const newDisData = _copyAndSort(
      tempDisArr,
      columnName,
      sortType == "desc" ? true : false
    );
    const newData = _copyAndSort(
      tempArr,
      columnName,
      sortType == "desc" ? true : false
    );

    setAdpData([...newData]);
    groups([...newDisData]);
  };

  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  const groups = (records) => {
    console.log(records, "ITS HRERE")
    let reOrderedRecords = [];

    let Uniquelessons = records.reduce(function (item, e1) {
      var matches = item.filter(function (e2) {
        return e1.Lesson === e2.Lesson;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    Uniquelessons.forEach((ul) => {
      let curLesson = records.filter((arr) => {
        return arr.Lesson == ul.Lesson;
      });
      reOrderedRecords = reOrderedRecords.concat(curLesson);
    });
    groupsforDL(reOrderedRecords);
  };

  const groupsforDL = (records) => {
    let newRecords = [];
    records.forEach((rd, index) => {
      newRecords.push({
        Lesson: rd.Lesson,
        indexValue: index,
      });
    });

    let varGroup = [];
    let Uniquelessons = newRecords.reduce(function (item, e1) {
      var matches = item.filter(function (e2) {
        return e1.Lesson === e2.Lesson;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    Uniquelessons.forEach((ul) => {
      let lessonLength = newRecords.filter((arr) => {
        return arr.Lesson == ul.Lesson;
      }).length;
      varGroup.push({
        key: ul.Lesson,
        name: ul.Lesson,
        startIndex: ul.indexValue,
        count: lessonLength,
      });
    });
    setAdpActivityResponseData([...records]);
    let filterRec = records.filter(item => {
      return item.Steps == 'Draft' ||
        item.Steps == 'Review' ||
        item.Steps == 'Edit' ||
        item.Steps == 'Assemble' ||
        item.Steps == 'Sign Off' ||
        item.Steps == 'Publish' ||
        item.Steps == 'Distribute';
    });
    formatData(records)
    setgroup([...varGroup]);
  };

  function checkIfLessonExists(lessonNumber: any, endResp: any) {
    for (let i = 0; i < endResp.length; i++) {
      if (endResp[i].LessonID === lessonNumber) {
        return true;
      }
    }
    return false;
  }

  const formater = (arrayData: any) => {
    const endResp: any = [];
    arrayData.map((x: any) => {
      if (!checkIfLessonExists(x.LessonID, endResp)) {
        endResp.push({
          LessonID: x.LessonID,
          l_name: x.Lesson,
          Project: x.Project,
          lesson: [
            { ...x }
            // [{ ...x }]
          ]
        })
      }
      else {
        const lessonIndex = endResp.findIndex((item: any) => item.LessonID === x.LessonID);
        let lessonIdAbs = endResp[lessonIndex];// main object
        const lessonData = endResp[lessonIndex].lesson; //main array


        lessonData.push({ ...x })
        endResp[lessonIndex] = {
          LessonID: lessonIdAbs.LessonID,
          l_name: lessonIdAbs.l_name,
          Project: x.Project,
          lesson: lessonData
        }
      }
    })



    return endResp;

  }


  //  :::: first add draft from step manually::::
  const revManager = (rev_loogs: any[]) => {

    const validResponses: { [key: string]: string } = {
      "Review": "Endorsed",
      "Initial Edit": "Edited",
      "Assemble": "Assembled",
      "Publish": "Publish ready",
      "Sign-off": "Signed Off",
      "Distribute": "Signed Off",

    };

    const resStatusChecker = (rev_item: any) => {
      if (rev_item.request == "Draft") {
        return true;
      }

      return validResponses[rev_item.request] === rev_item.response;
    };

    var finalArray: any[] = []
    let currentlistedArray = 0;
    let count = 1;
    let req_type = "";
    let res_type: boolean = false;


    rev_loogs.forEach(function (rev_item: any) {

      if (count == 1) {
        finalArray[currentlistedArray] = [{ ...rev_item, count_arr: count }]
      } else if (res_type === true) {
        finalArray[currentlistedArray].push({ ...rev_item, count_arr: count })
      }
      else {
        finalArray[currentlistedArray + 1] = [{ ...rev_item, count_arr: count }]
        currentlistedArray = currentlistedArray + 1;
      }
      count = count + 1;
      req_type = rev_item.request;
      res_type = resStatusChecker(rev_item);
    });

    return finalArray;
    console.log(finalArray, "finalArray")
  }



  //Function-Section Ends
  useEffect(() => {
    if (
      adpAutoSave &&
      adpEditFlag &&
      adpData.some((data) => data.dateError == true) == false
    ) {
      setTimeout(() => {
        newDataFlag
          ? document.getElementById("adpbtnSave").click()
          : document.getElementById("adpbtnUpdate").click();
      }, 300000);
    }
  }, [adpAutoSave]);

  useEffect(() => {
    setAdpLoader("startUpLoader");
    setAdpLoader2("startUpLoader");
    getActivityPlanItem();
    getActivityPBData();
    adpGetCurrentUserDetails();
  }, [adpReRender]);
  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {adpLoader2 == "startUpLoader" ? <CustomLoader /> : null}
        {/* Header-Section Starts */}
        <div className={styles.adpHeaderSection} style={{ paddingBottom: "0" }}>
          {/* Popup-Section Starts */}
          <div></div>
          {/* Popup-Section Ends */}
          <div className={styles.adpHeader} style={{ marginBottom: "15px" }}>
            <div className={styles.dpTitle}>
              <Icon
                iconName="NavigateBack"
                className={adpIconStyleClass.navArrow}
                onClick={() => {
                  adpAutoSave
                    ? confirm(
                      "You have unsaved changes, are you sure you want to leave?"
                    )
                      ? props.handleclick("ActivityPlan", null, "adp")
                      : null
                    : props.handleclick("ActivityPlan", null, "adp");
                }}
              />
              <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
                Activity planner
              </Label>
            </div>
            {/* <div style={{ display: "flex" }}>
              <Persona
                size={PersonaSize.size32}
                presence={PersonaPresence.none}
                imageUrl={
                  "/_layouts/15/userphoto.aspx?size=S&username=" +
                  `${
                    activtyPlanItem.length > 0
                      ? activtyPlanItem[0]["DeveloperDetails"].email
                      : ""
                  }`
                }
              />
              <Label>
                {activtyPlanItem.length > 0
                  ? activtyPlanItem[0]["DeveloperDetails"].name
                  : ""}
              </Label>
            </div> */}
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              flexWrap: "wrap",
            }}
          >
            <div
              className={styles.adpHeaderDetails}
              style={{ marginLeft: "-10px" }}
            >
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0
                    ? activtyPlanItem[0].Project +
                    " " +
                    activtyPlanItem[0].ProjectVersion
                    : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Product :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0
                    ? activtyPlanItem[0].Product +
                    " " +
                    activtyPlanItem[0].ProductVersion
                    : ""}
                </Label>
              </div>
              <div>
                {/* <Label style={{ marginRight: 5 }}>
                  Number of records :{" "}
                  <span style={{ color: "#038387" }}>
                    {adpActivityResponseData.length}
                  </span>
                </Label> */}
              </div>
              {/* <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Status :</Label>
                <Label style={{ color: "#038387", marginRight: "-25px" }}>
                  {overallStatus()}
                </Label>
              </div> *
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Type :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>AH/PH :</Label>
                <Label style={{ color: "#038387" }}>
                  {overallActualHours()}/{overallPlannedHours()}
                </Label>
              </div> */}
            </div>

          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "space-between",
              flexWrap: "wrap",
            }}
          >
            <div
              className={styles.adpHeaderDetails}
              style={{ marginLeft: "-10px" }}
            >
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>BA :</Label>
                <Label style={{ color: "#038387" }}>
                  {annualPlanData.ba}

                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Term :</Label>
                <Label style={{ color: "#038387" }}>
                  {annualPlanData.term}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Hours :</Label>
                <Label style={{ color: "#038387" }}> {totalHours}
                  {/* {activtyPlanItem.length > 0
                    ? activtyPlanItem[0].Project :''
                   } */}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Start :</Label>
                <Label style={{ color: "#038387" }}>
                  {Tabledata.length > 0
                    ? Tabledata[0][0].Start : ''
                  }
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>End :</Label>
                <Label style={{ color: "#038387" }}>
                  {Tabledata.length > 0
                    ? Tabledata[Tabledata.length - 1][Tabledata[Tabledata.length - 1].length - 1].End
                    : ''}
                </Label>
              </div>

              {/* <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Status :</Label>
                <Label style={{ color: "#038387", marginRight: "-25px" }}>
                  {overallStatus()}
                </Label>
              </div> *
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Type :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>Project :</Label>
                <Label style={{ color: "#038387" }}>
                  {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
                </Label>
              </div>
              <div style={{ margin: "0 25px 0 10px" }}>
                <Label style={{ marginRight: 5 }}>AH/PH :</Label>
                <Label style={{ color: "#038387" }}>
                  {overallActualHours()}/{overallPlannedHours()}
                </Label>
              </div> */}
            </div>
            <div style={{ display: "flex" }}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "flex-end",
                  marginTop: 2,
                  marginRight: 20,
                }}
              >

              </div>
              {/* <Label
                onClick={() => {
                  generateExcel();
                }}
                style={{
                  backgroundColor: "#EBEBEB",
                  padding: "0 15px",
                  cursor: "pointer",
                  fontSize: "12px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "3px",
                  color: "#1D6F42",
                  height: 34,
                  marginRight: 10,
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                    marginRight: 5,
                  }}
                  iconName="ExcelDocument"
                  className={adpIconStyleClass.export}
                />
                Export as XLS
              </Label> */}
              {adpEditFlag ? (
                <PrimaryButton
                  className={adpbuttonStyleClass.buttonPrimary}
                  iconProps={cancelIcon}
                  onClick={() => {
                    setAdpEDSort("");
                    setAdpSDSort("");
                    setAdpEditFlag(false);
                    setAdpAutoSave(false);
                    // setAdpData([...adpMasterData]);
                    // groups([...adpMasterData]);
                    // setAdpActivityResponseData([...adpMasterData]);
                    // setFinalStepConst(datauniqueArr)
                    setFinalStep(finalStepConst)
                    // setAdpFilters({ ...adpFilterKeys });
                  }}
                >
                  Cancel
                </PrimaryButton>
              ) : (
                <PrimaryButton
                  className={adpbuttonStyleClass.buttonPrimary}
                  iconProps={editIcon}
                  onClick={() => {
                    setAdpEditFlag(true);
                    setAdpAutoSave(true);
                  }}
                >
                  Edit
                </PrimaryButton>
              )}
              {/* {newDataFlag == true && adpItemAddFlag == false ? (
                <PrimaryButton
                  id="adpbtnSave"
                  iconProps={saveIcon}
                  className={
                    adpEditFlag &&
                      adpData.some(
                        (data) => data.dateError == true || data.PHError == true
                      ) == false
                      ? adpbuttonStyleClass.buttonSecondary
                      : styles.adpSaveBtnDisabled
                  }
                  disabled={
                    adpEditFlag &&
                      adpData.some(
                        (data) => data.dateError == true || data.PHError == true
                      ) == false
                      ? false
                      : true
                  }
                  onClick={() => {
                    console.log(diliveryPlanNeedToBeUpdate, "updated h 1???")
                    adpUpdateItem(1)
                    // if (adpEditFlag) {
                    //   setAdpAutoSave(false);

                    //   let isCompletedData = adpData.filter((arr) => {
                    //     return arr.IsCompleteNew == true;
                    //   });
                    //   if (isCompletedData.length > 0) {
                    //     setAdpConfirmationPopup({
                    //       condition: true,
                    //       isNew: true,
                    //     });
                    //   } else {
                    //     setAdpLoader("startUpLoader");
                    //     adpAddItem();
                    //   }
                    //   // setAdpLoader("startUpLoader");
                    //   // adpAddItem();
                    // }
                  }}
                >
                  {adpLoader == "saveLoader" ? <Spinner /> : <>Save</>}
                </PrimaryButton>
              ) : ( */}
              <PrimaryButton
                id="adpbtnUpdate"
                iconProps={saveIcon}
                className={
                  adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                    ? adpbuttonStyleClass.buttonSecondary
                    : styles.adpSaveBtnDisabled
                }
                disabled={
                  adpEditFlag &&
                    adpData.some(
                      (data) => data.dateError == true || data.PHError == true
                    ) == false
                    ? false
                    : true
                }
                onClick={() => {
                  console.log(diliveryPlanNeedToBeUpdate, "updated h???")
                  adpUpdateItem()
                  // if (
                  //   !adpData.some(
                  //     (data) => data.dateError == true || data.PHError == true
                  //   )
                  // ) {
                  //   if (adpEditFlag) {
                  //     setAdpAutoSave(false);

                  //     let isCompletedData = adpData.filter((arr) => {
                  //       return arr.IsCompleteNew == true;
                  //     });
                  //     if (isCompletedData.length > 0) {
                  //       setAdpConfirmationPopup({
                  //         condition: true,
                  //         isNew: false,
                  //       });
                  //     } else {
                  //       setAdpLoader("startUpLoader");
                  //       // adpUpdateItem();
                  //     }

                  //     // setAdpLoader("startUpLoader");
                  //     // adpUpdateItem();
                  //   }
                  // }
                }}
              >
                {adpLoader == "updateLoader" ? <Spinner /> : <>Save</>}
              </PrimaryButton>
              {/* )} */}
              <Icon
                iconName="Link12"
                className={adpIconStyleClass.link}
                onClick={() => {
                  adpAutoSave
                    ? confirm(
                      "You have unsaved changes, are you sure you want to leave?"
                    )
                      ? props.handleclick(
                        "ActivityProductionBoard",
                        activityPlan_ID,
                        "ADP"
                      )
                      : null
                    : props.handleclick(
                      "ActivityProductionBoard",
                      activityPlan_ID,
                      "ADP"
                    );
                }}
              />
            </div>
          </div>
          {/* Header-Section Ends */}
          {/* Filter-Section Starts */}
          <div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                marginTop: "-5px",
                marginBottom: "10px",
                flexWrap: "wrap",
              }}
            >
              {/* <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexWrap: "wrap",
                }}
              >
              
                <div>
                  <Label styles={adpLabelStyles}>Section</Label>
                  <SearchBox
                    styles={
                      adpFilters.lesson
                        ? adpActiveSearchBoxStyles
                        : adpSearchBoxStyles
                    }
                    value={adpFilters.lesson}
                    onChange={(e, value) => {
                      adpListFilter("lesson", value);
                    }}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Steps</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.step != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.stepsOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("step", option["key"]);
                    }}
                    selectedKey={adpFilters.step}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Status</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.status != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.statusOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("status", option["key"]);
                    }}
                    selectedKey={adpFilters.status}
                  />
                </div>
                <div>
                  <Label styles={adpLabelStyles}>Developer</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.developer != "All"
                        ? adpActiveDropdownStyles
                        : adpDropdownStyles
                    }
                    options={adpDropDownOptions.developerOptns}
                    dropdownWidth={"auto"}
                    onChange={(e, option: any) => {
                      adpListFilter("developer", option["key"]);
                    }}
                    selectedKey={adpFilters.developer}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      width: 75,
                    }}
                    styles={adpLabelStyles}
                  >
                    Year
                  </Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.year != "All"
                        ? adpActiveShortDropdownStyles
                        : adpShortDropdownStyles
                    }
                    options={adpDropDownOptions.yearOptns}
                    onChange={(e, option: any) => {
                      adpListFilter("year", option["key"]);
                    }}
                    selectedKey={adpFilters.year}
                  />
                </div>
                <div>
                  <Label
                    style={{
                      width: 75,
                    }}
                    styles={adpLabelStyles}
                  >
                    Week
                  </Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      adpFilters.week != "All"
                        ? adpActiveShortDropdownStyles
                        : adpShortDropdownStyles
                    }
                    options={adpDropDownOptions.weekOptns}
                    onChange={(e, option: any) => {
                      adpListFilter("week", option["key"]);
                    }}
                    selectedKey={adpFilters.week}
                  />
                </div>

                <div>
                  <Label
                    style={{
                      width: 65,
                    }}
                    styles={adpLabelStyles}
                  >
                    Start date
                  </Label>
                  <div
                    style={{
                      display: "flex",
                      marginTop: 5,
                      marginRight: 15,
                    }}
                  >
                    <button
                      style={{
                        backgroundColor: "#038387",
                        border: 0,
                        borderRadius: 10,
                        padding: "0.25rem  0.5rem",
                        marginRight: 10,
                        cursor: "pointer",
                      }}
                    >
                      <Icon
                        title={"asc"}
                        style={{
                          color: "#fff",
                          fontSize: adpSDSort == "asc" ? 20 : 16,
                          fontWeight: adpSDSort == "asc" ? "bold" : "normal",
                        }}
                        iconName="SortUp"
                        onClick={() => {
                          setAdpEDSort("");
                          setAdpSDSort("asc");
                          sortingFunction("Start", "asc");
                        }}
                      />
                      <Icon
                        title={"desc"}
                        style={{
                          color: "#fff",
                          fontSize: adpSDSort == "desc" ? 20 : 16,
                          fontWeight: adpSDSort == "desc" ? "bold" : "normal",
                        }}
                        iconName="SortDown"
                        onClick={() => {
                          setAdpEDSort("");
                          setAdpSDSort("desc");
                          sortingFunction("Start", "desc");
                        }}
                      />
                    </button>
                  </div>
                </div>
                <div>
                  <Label
                    style={{
                      width: 65,
                    }}
                    styles={adpLabelStyles}
                  >
                    End date
                  </Label>
                  <div
                    style={{
                      display: "flex",
                      marginTop: 5,
                      marginRight: 15,
                    }}
                  >
                    <button
                      style={{
                        backgroundColor: "#038387",
                        border: 0,
                        borderRadius: 10,
                        padding: "0.25rem  0.5rem",
                        marginRight: 10,
                        cursor: "pointer",
                      }}
                    >
                      <Icon
                        title={"asc"}
                        style={{
                          color: "#fff",
                          fontSize: adpEDSort == "asc" ? 20 : 16,
                          fontWeight: adpEDSort == "asc" ? "bold" : "normal",
                        }}
                        iconName="SortUp"
                        onClick={() => {
                          setAdpSDSort("");
                          setAdpEDSort("asc");
                          sortingFunction("End", "asc");
                        }}
                      />
                      <Icon
                        title={"desc"}
                        style={{
                          color: "#fff",
                          fontSize: adpEDSort == "desc" ? 20 : 16,
                          fontWeight: adpEDSort == "desc" ? "bold" : "normal",
                        }}
                        iconName="SortDown"
                        onClick={() => {
                          setAdpSDSort("");
                          setAdpEDSort("desc");
                          sortingFunction("End", "desc");
                        }}
                      />
                    </button>
                  </div>
                </div>
                <div>
                  <Label style={{ width: 60 }} styles={adpLabelStyles}>
                    Complete
                  </Label>
                  <Checkbox
                    styles={{
                      root: { marginTop: 3, width: 50 },
                    }}
                    disabled={!adpEditFlag ? true : false}
                    checked={AdpIsCompleted}
                    onChange={(ev) => {
                      setAdpIsCompleted(!AdpIsCompleted);
                      adpData.forEach((item, Index) => {
                        let dpBeforeData = adpData[Index];
                        let dpOnchangeData = [
                          {
                            OrderId: dpBeforeData.OrderId,
                            ID: dpBeforeData.ID,
                            Steps: dpBeforeData.Steps,
                            PH: dpBeforeData.PH,
                            MinPH: dpBeforeData.MinPH,
                            MaxPH: dpBeforeData.MaxPH,
                            Project: dpBeforeData.Project,
                            LessonID: dpBeforeData.LessonID,
                            IsCompleteStatus:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteStatus,
                            Lesson: dpBeforeData.Lesson,
                            Start: dpBeforeData.Start,
                            End: dpBeforeData.End,
                            Developer: dpBeforeData.Developer,
                            Status: dpBeforeData.Status,
                            AH: dpBeforeData.AH,
                            dateError: dpBeforeData.dateError,
                            PHError: dpBeforeData.PHError,
                            PHWeek: dpBeforeData.PHWeek,
                            IsCompleteNew:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteNew,
                          },
                        ];
                        adpData[Index] = dpOnchangeData[0];
                      });

                      adpActivityResponseData.forEach((item, Index) => {
                        let dpBeforeData = adpActivityResponseData[Index];
                        let dpOnchangeData = [
                          {
                            OrderId: dpBeforeData.OrderId,
                            ID: dpBeforeData.ID,
                            Steps: dpBeforeData.Steps,
                            PH: dpBeforeData.PH,
                            MinPH: dpBeforeData.MinPH,
                            MaxPH: dpBeforeData.MaxPH,
                            Project: dpBeforeData.Project,
                            LessonID: dpBeforeData.LessonID,
                            IsCompleteStatus:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteStatus,
                            Lesson: dpBeforeData.Lesson,
                            Start: dpBeforeData.Start,
                            End: dpBeforeData.End,
                            Developer: dpBeforeData.Developer,
                            Status: dpBeforeData.Status,
                            AH: dpBeforeData.AH,
                            dateError: dpBeforeData.dateError,
                            PHError: dpBeforeData.PHError,
                            PHWeek: dpBeforeData.PHWeek,
                            IsCompleteNew:
                              item.Status != "Completed"
                                ? ev.target["checked"]
                                : dpBeforeData.IsCompleteNew,
                          },
                        ];
                        adpActivityResponseData[Index] = dpOnchangeData[0];
                      });

                      setAdpData([...adpData]);
                      setAdpActivityResponseData([...adpActivityResponseData]);
                    }}
                  />
                </div>
                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={adpIconStyleClass.refresh}
                    onClick={() => {
                      if (adpAutoSave) {
                        if (
                          confirm(
                            "You have unsaved changes, are you sure you want to leave?"
                          )
                        ) {
                          setAdpEDSort("");
                          setAdpSDSort("");
                          groups(adpMasterData);
                          // setAdpActivityResponseData(adpMasterData);
                          setAdpData([...adpMasterData]);
                          adpGetAllOptions(adpMasterData);
                          setAdpFilters({ ...adpFilterKeys });
                        }
                      } else {
                        setAdpEDSort("");
                        setAdpSDSort("");
                        groups(adpMasterData);
                        // setAdpActivityResponseData(adpMasterData);
                        setAdpData([...adpMasterData]);
                        adpGetAllOptions(adpMasterData);
                        setAdpFilters({ ...adpFilterKeys });
                      }
                    }}
                  />
                </div>
                <div>
                  <div
                    style={{
                      // display: "flex",
                      // justifyContent: "flex-end",
                      // marginLeft: "20px",
                      marginTop: "38px",
                    }}
                  >
                    {adpEditFlag &&
                      adpData.some((data) => data.dateError == true) ? (
                      <Label
                        style={{
                          marginRight: 5,
                        }}
                        className={adpCommonStyles.dateGridValidationErrorLabel}
                      >
                        *Given end date should not be earlier than the start
                        date
                      </Label>
                    ) : null}
                    {adpEditFlag &&
                      adpData.some((data) => data.PHError == true) ? (
                      <Label
                        style={{
                          marginRight: 5,
                        }}
                        className={adpCommonStyles.dateGridValidationErrorLabel}
                      >
                        *Please enter valid hours(PH)
                      </Label>
                    ) : null}
                  </div>
                </div>
              </div> */}
            </div>
          </div>
          {/* Filter-Section Ends */}
        </div>

        {/* Body-Section Starts */}

        <div>
          {/* dont remove */}
          {/* <input
            id="forFocus"
            type="text"
            style={{
              width: 0,
              height: 0,
              border: "none",
              position: "absolute",
              top: 0,
              left: 0,
              padding: 0,
            }}
          /> */}
        </div>
        <div
          className={styles.scrollTop}
          onClick={() => {
            document.querySelector("#forFocus")["focus"]();
          }}
        >
          <Icon iconName="Up" style={{ color: "#fff" }} />
        </div>
        <div>
          {/* Table-Section Starts */}




          {/* <div className="lessonDiv">
            {Open?.some((arrVal) => index == arrVal) ? (
              <svg xmlns="http://www.w3.org/2000/svg" style={{
                width: '54px',

                height: '16px', cursor: 'pointer'
              }} viewBox="0 0 24 24" onClick={() => show(index)}>
                <path d="M7 14l5-5 5 5z" />
              </svg>
            ) : (
              <svg xmlns="http://www.w3.org/2000/svg" style={{
                width: '24px',

                height: '16px', cursor: 'pointer'
              }} viewBox="0 0 24 24" onClick={() => show(index)}>
                <path d="M12 14l5-5H7z" />
              </svg>)}


              Lesson {data[0]?.LessonID + '(' + data.length + ')' + ' '}
               {((data.filter(task => task.Status === 'Completed').length / data.length) * 100)+"% Complete"}
            </div> */}
          <div style={{ overflowX: "auto" }} id="style-1" className="MyTable_2">
            {/* {Open?.some((arrVal) => index == arrVal) && <>  */}
            <table className="table table-bordered fixed-width-table">

              <tr style={{ position: "sticky", top: "0", zIndex: "6" }}>
                <th className="Title text-center-do" style={{ width: "16%" }} >Task</th>
                <th className="Title text-center-do" style={{ width: "12%" }}>Activity</th>



                {StepsArrayCustomized.map
                  ((stepInfo, i) =>
                  (
                    <th className="Title  text-center-do" style={{ width: "12%" }} colSpan={2}>{stepInfo.name}
                      {/* // + '(' + stepInfo.ph + ' hrs)'} */}
                    </th>))
                }


              </tr>
              {/* {

                objRespTemplate.map((data, ind) => {
                  return (<>
                    <tr>
                      <td className="typeData" width="15%"> {ind == 0 && data.task}</td>
                      <td width="13%">{ind == 0 && data.Lesson}</td>

                      <td style={{
                      }} width="12%">
                        {data.step}
                      </td>
                      <td style={{
                      }}>

                        13/06/2023
                      </td>
                      <td

                      >

                        15/06/2023
                      </td>
                      <td
                      >

                        16/06/2023
                      </td>

                      <td  >

                        18/06/2023
                      </td>
                      <td  >


                        20/06/2023
                      </td>

                      <td  >

                        25/06/2023


                      </td>
                      <td  >
                        26/06/2023
                      </td>
                      <td  >

                        27/07/2023
                      </td>
                      <td  >
                        28/07/2022

                      </td>
                      <td  >

                        29/06/2023
                      </td>
                      <td  >
                        30/06/2023
                      </td>

                      <td  >

                        31/06/2023
                      </td>




                    </tr>
                    <tr>
                      <td className="typeData" width="15%">  </td>
                      <td width="13%"> </td>

                      <td style={{
                      }} width="12%">

                      </td>
                      <td style={{
                      }}>


                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.d_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>
                      <td

                      >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.d_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>

                      </td>
                      <td
                      >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.rev_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>

                      </td>

                      <td  >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.rev_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>
                      <td  >

                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.edit_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>

                      </td>

                      <td  >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.edit_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>



                      </td>
                      <td  >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.ass_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>
                      <td  >

                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.ass_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>
                      <td  >
                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.ass_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>

                      </td>
                      <td  >

                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.ass_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>
                      <td  >

                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.aprv_dev_1}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>

                      <td  >

                        <>
                          <TooltipHost
                            content={"dev name"}
                            id="myPersonaTooltip"
                          >
                            <Persona
                              size={PersonaSize.size32}
                              presence={PersonaPresence.none}
                              imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${data.aprv_dev_2}`
                              }
                            />
                          </TooltipHost>
                        </>
                      </td>




                    </tr>
                  </>)
                })

              } */}



              {finalStep.length > 0 ? finalStep.map((data, index) => {

                return <>

                  {data.lesson.map((nestedArrays, ind) => {
                    return <>

                      <tr>
                        <td className="typeData" > {ind == 0 && data.Project}</td>
                        <td >{ind == 0 && data.l_name}</td>

                        <>


                          <td>
                            {/* {nestedArrays.find((obj) => obj.request === "Draft") ? nestedArrays.find((obj) => obj.request === "Draft").Start : ''} */}


                            {ind == 0 && adpEditFlag && nestedArrays.find((obj) => obj.request === "Draft") ? (
                              <>
                                <DatePicker
                                  placeholder="Select a start date"
                                  formatDate={dateFormater}
                                  // minDate={new Date(item.Start)}
                                  // maxDate={new Date(item.End)}
                                  styles={{
                                    textField: {
                                      transform: "translateY(3px)",
                                      selectors: {
                                        ".ms-TextField-fieldGroup": {
                                          borderColor: "#000",
                                          borderRadius: 4,
                                          border: "1px solid",
                                          height: 23,
                                          input: {
                                            borderRadius: 4,
                                          },
                                        },
                                        ".ms-TextField-field": {
                                          color: "#000",
                                        },
                                        ".ms-DatePicker-event--without-label": {
                                          color: "#000",
                                          paddingTop: 3,
                                        },
                                      },
                                    },
                                    readOnlyTextField: {
                                      lineHeight: 22,
                                    },
                                  }}
                                  value={
                                    nestedArrays.find((obj) => obj.request === "Draft")
                                      ? new Date(
                                        moment(nestedArrays.find((obj) => obj.request === "Draft").Start, DateListFormat).format(DatePickerFormat)
                                      )
                                      : new Date()
                                  }
                                  onSelectDate={(value: any) => {
                                    console.log("----")

                                    // const inputDate = new Date(value);
                                    // const convertedDate = convertDateFormat(inputDate);
                                    // console.log(convertedDate);




                                    const inputDate = new Date(value);
                                    const customFormat = "DD/MM/YYYY";
                                    const convertedDate = convertToCustomFormat(inputDate, customFormat);

                                    console.log("change date", convertedDate)
                                    console.log("Draft", nestedArrays.find((obj) => obj.request === "Draft"))
                                    const selectedObj = nestedArrays.findIndex((obj) => obj.request === "Draft")

                                    let finalStepTemp = finalStep;

                                    let respData = finalStepTemp[index].lesson[ind].find((obj) => obj.request === "Draft")

                                    if (convertedDate) {
                                      respData.Start = convertedDate
                                    }

                                    const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
                                    const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.ADPId === respData.ADPId)
                                    if (dilIndex == -1) {
                                      diliveryPlanNeedToBeUpdateTemp.push(respData)
                                    } else {
                                      diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
                                    }
                                    setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update
                                    finalStepTemp[index].lesson[ind][selectedObj] = respData
                                    console.log(ind, "nested index")
                                    console.log(respData, "after")
                                    console.log(finalStepTemp, "1234567890")
                                    setFinalStep(finalStepTemp)
                                    setreRenderState(!reRenderState)


                                    // adpActivityResponseHandler(item.OrderId, "End", value);
                                  }}
                                />
                              </>
                            ) : (
                              <>
                                {ind == 0 &&
                                  nestedArrays.find((obj) => obj.request === "Draft") && (
                                    <>{nestedArrays.find((obj) => obj.request === "Draft").Start}</>
                                  )
                                }
                              </>
                            )}
                          </td>
                          <td  >
                            {/* {nestedArrays.find((obj) => obj.request === "Draft") ? nestedArrays.find((obj) => obj.request === "Draft").End : ''} */}
                            {ind == 0 && adpEditFlag && nestedArrays.find((obj) => obj.request === "Draft") ? (
                              <>
                                <DatePicker
                                  placeholder="Select a start date"
                                  formatDate={dateFormater}
                                  // minDate={new Date(item.Start)}
                                  // maxDate={new Date(item.End)}
                                  styles={{
                                    textField: {
                                      transform: "translateY(3px)",
                                      selectors: {
                                        ".ms-TextField-fieldGroup": {
                                          borderColor: "#000",
                                          borderRadius: 4,
                                          border: "1px solid",
                                          height: 23,
                                          input: {
                                            borderRadius: 4,
                                          },
                                        },
                                        ".ms-TextField-field": {
                                          color: "#000",
                                        },
                                        ".ms-DatePicker-event--without-label": {
                                          color: "#000",
                                          paddingTop: 3,
                                        },
                                      },
                                    },
                                    readOnlyTextField: {
                                      lineHeight: 22,
                                    },
                                  }}
                                  value={
                                    nestedArrays.find((obj) => obj.request === "Draft")
                                      ? new Date(
                                        moment(nestedArrays.find((obj) => obj.request === "Draft").End, DateListFormat).format(DatePickerFormat)
                                      )
                                      : new Date()
                                  }
                                  onSelectDate={(value: any) => {
                                    console.log("----")
                                    // adpActivityResponseHandler(item.OrderId, "End", value);




                                    const inputDate = new Date(value);
                                    const customFormat = "DD/MM/YYYY";
                                    const convertedDate = convertToCustomFormat(inputDate, customFormat);

                                    console.log("change date", convertedDate)
                                    console.log("Draft", nestedArrays.find((obj) => obj.request === "Draft"))
                                    const selectedObj = nestedArrays.findIndex((obj) => obj.request === "Draft")

                                    let finalStepTemp = finalStep;

                                    let respData = finalStepTemp[index].lesson[ind].find((obj) => obj.request === "Draft")

                                    if (convertedDate) {
                                      respData.End = convertedDate
                                    }

                                    const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
                                    const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.ADPId === respData.ADPId)
                                    if (dilIndex == -1) {
                                      diliveryPlanNeedToBeUpdateTemp.push(respData)
                                    } else {
                                      diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
                                    }
                                    setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update
                                    finalStepTemp[index].lesson[ind][selectedObj] = respData;
                                    console.log(ind, "nested index")
                                    console.log(respData, "after")
                                    console.log(finalStepTemp, "1234567890")
                                    setFinalStep(finalStepTemp)
                                    setreRenderState(!reRenderState)


                                  }}
                                />
                              </>
                            ) : (
                              <>
                                {ind == 0 &&
                                  nestedArrays.find((obj) => obj.request === "Draft") && (
                                    <>{nestedArrays.find((obj) => obj.request === "Draft").End}</>
                                  )
                                }
                              </>
                            )}
                          </td>
                          <td

                          >
                            {nestedArrays.find((obj) => obj.request === "Review") ? dateFormater(nestedArrays.find((obj) => obj.request === "Review").Start) : ''}

                          </td>
                          <td
                          >
                            {nestedArrays.find((obj) => obj.request === "Review") ? dateFormater(nestedArrays.find((obj) => obj.request === "Review").End) : ''}

                          </td>

                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Initial Edit") ? dateFormater(nestedArrays.find((obj) => obj.request === "Initial Edit").Start) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Initial Edit") ? dateFormater(nestedArrays.find((obj) => obj.request === "Initial Edit").End) : ''}

                          </td>

                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Assemble") ? dateFormater(nestedArrays.find((obj) => obj.request === "Assemble").Start) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Assemble") ? dateFormater(nestedArrays.find((obj) => obj.request === "Assemble").End) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Sign-off") ? nestedArrays.find((obj) => obj.request === "Sign-off").Start : nestedArrays.find((obj) => obj.request === "Publish") ? dateFormater(nestedArrays.find((obj) => obj.request === "Publish").Start) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Sign-off") ? nestedArrays.find((obj) => obj.request === "Sign-off").Start : nestedArrays.find((obj) => obj.request === "Publish") ? dateFormater(nestedArrays.find((obj) => obj.request === "Publish").Start) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Distribute") ? dateFormater(nestedArrays.find((obj) => obj.request === "Distribute").Start) : ''}

                          </td>
                          <td  >
                            {nestedArrays.find((obj) => obj.request === "Distribute") ? dateFormater(nestedArrays.find((obj) => obj.request === "Distribute").End) : ''}

                          </td>


                        </>


                      </tr>

                      <tr>
                        <td className="typeData"  >
                        </td>
                        <td  > </td>

                        <>
                          <td colSpan={2}  >




                            {ind == 0 &&
                              adpEditFlag ? (
                              <>
                                <NormalPeoplePicker
                                  styles={{
                                    root: {
                                      selectors: {
                                        ".ms-SelectionZone": {
                                          height: 24,
                                        },
                                        ".ms-BasePicker-text": {
                                          height: 24,
                                          padding: 1,
                                          border: "1px solid #000",
                                          borderRadius: 4,
                                          marginTop: -6,
                                          marginRight: 20,
                                        },
                                      },
                                    },
                                  }}
                                  onResolveSuggestions={GetUserDetails}
                                  itemLimit={1}
                                  selectedItems={allPeoples.filter((people) => {
                                    return (
                                      people.ID == (nestedArrays.find((obj) => obj?.request === "Draft")?.Dev2?.id ? nestedArrays.find((obj) => obj?.request === "Draft")?.Dev2?.id : null)
                                    );
                                  })}
                                  onChange={(selectedUser) => {
                                    console.log("change image", selectedUser)
                                    console.log("Draft", nestedArrays.find((obj) => obj.request === "Draft"))
                                    const selectedObj = nestedArrays.findIndex((obj) => obj.request === "Draft")
                                    console.log(finalStep, "masti kr rya h")
                                    let finalStepTemp = finalStep;

                                    let respData = finalStepTemp[index].lesson[ind].find((obj) => obj.request === "Draft")
                                    console.log(respData, "before")
                                    if (selectedUser[0]) {


                                      respData.Dev2 = {
                                        id: selectedUser[0]["ID"],
                                        email: selectedUser[0]["secondaryText"],
                                        name: selectedUser[0]["text"]
                                      }
                                      respData.Dev = selectedUser[0]["text"]
                                      respData.FromEmail = selectedUser[0]["secondaryText"]
                                      respData.ToEmail = selectedUser[0]["secondaryText"]
                                    } else {



                                      respData.Dev2 = {
                                        id: undefined,
                                        email: undefined,
                                        name: undefined
                                      }
                                      respData.Dev = undefined
                                      respData.FromEmail = undefined
                                      respData.ToEmail = undefined

                                    }

                                    const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
                                    const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.ADPId === respData.ADPId)
                                    if (dilIndex == -1) {
                                      diliveryPlanNeedToBeUpdateTemp.push(respData)
                                    } else {
                                      diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
                                    }
                                    setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update
                                    finalStepTemp[index].lesson[ind][selectedObj] = respData
                                    console.log(ind, "nested index")
                                    console.log(respData, "after")
                                    console.log(finalStepTemp, "1234567890")
                                    setFinalStep(finalStepTemp)
                                    setreRenderState(!reRenderState)



                                    // adpActivityResponseHandler(
                                    //   1,
                                    //   "Developer",
                                    //   selectedUser[0] ? selectedUser[0]["ID"] : null
                                    // ); 
                                  }}
                                />
                              </>
                            ) : (
                              <>
                                {ind == 0 && newDataFlag ? (
                                  <>
                                    <TooltipHost
                                      id="myPersonaTooltip"
                                      style={{ display: "flex", justifyContent: "center" }}
                                    >
                                      <Persona
                                        size={PersonaSize.size32}
                                        presence={PersonaPresence.none}
                                        imageUrl={
                                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                                          `${nestedArrays.find((obj) => obj?.request === "Draft")?.FromEmail}`
                                        }
                                      />
                                    </TooltipHost>

                                  </>
                                ) : ind == 0 && nestedArrays.find((obj) => obj?.request === "Draft")?.FromEmail ? (
                                  <>
                                    <TooltipHost
                                      id="myPersonaTooltip"
                                      style={{ display: "flex", justifyContent: "center" }}
                                    >
                                      <Persona
                                        size={PersonaSize.size32}
                                        presence={PersonaPresence.none}
                                        imageUrl={
                                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                                          `${nestedArrays.find((obj) => obj?.request === "Draft")?.FromEmail}`
                                        }
                                      />
                                    </TooltipHost>

                                  </>
                                ) : (
                                  ""
                                )}
                              </>
                            )
                            }



                            {/* 
                            {nestedArrays.find((obj) => obj?.request === "Draft") ? 
                              <>
                                <TooltipHost
                                  id="myPersonaTooltip"
                                  style={{ display: "flex", justifyContent: "center" }}
                                >
                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Draft")?.FromEmail}`
                                    }
                                  />
                                </TooltipHost>

                              </>
 
                              : ''}

*/}


                          </td>

                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Review") ?

                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Review")?.Dev}
                                  id="myPersonaTooltip"
                                >
                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Review")?.FromEmail}`
                                    }
                                  />
                                </TooltipHost>

                              </>


                              : ''}

                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Review") ?
                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Review")?.client}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Review")?.ToEmail}`
                                    }
                                  />

                                </TooltipHost>
                              </>


                              : ''}

                          </td>


                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Initial Edit") ?
                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Initial Edit")?.Dev}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Initial Edit")?.FromEmail}`
                                    }
                                  />
                                </TooltipHost>
                              </>



                              : ''}

                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Initial Edit") ?

                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Initial Edit")?.client}
                                  id="myPersonaTooltip"
                                >
                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Initial Edit")?.ToEmail}`
                                    }
                                  /></TooltipHost></>




                              : ''}

                          </td>



                          <td  >


                            {nestedArrays.find((obj) => obj?.request === "Assemble") ?
                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Assemble")?.Dev}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Assemble")?.FromEmail}`
                                    }
                                  />
                                </TooltipHost></>



                              : ''}

                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Assemble") ?

                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Assemble")?.client}
                                  id="myPersonaTooltip"
                                >
                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Assemble")?.ToEmail}`
                                    }
                                  />
                                </TooltipHost></>



                              : ''}

                          </td>       <td  >

                            {nestedArrays.find((obj) => obj?.request === "Sign-off") ?

                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Sign-off")?.Dev}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Sign-off")?.FromEmail}`
                                    }
                                  /> </TooltipHost> </> : nestedArrays.find((obj) => obj.request === "Publish") ?
                                <>
                                  <TooltipHost
                                    content={nestedArrays.find((obj) => obj?.request === "Publish")?.Dev}
                                    id="myPersonaTooltip"
                                  >

                                    <Persona
                                      size={PersonaSize.size32}
                                      presence={PersonaPresence.none}
                                      imageUrl={
                                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                                        `${nestedArrays.find((obj) => obj?.request === "Publish")?.FromEmail}`
                                      }
                                    /></TooltipHost> </> : ''}

                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Sign-off") ?
                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Sign-off")?.client}
                                  id="myPersonaTooltip"
                                >
                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Sign-off")?.ToEmail}`
                                    }
                                  /> </TooltipHost>
                              </> : nestedArrays.find((obj) => obj.request === "Publish") ?

                                <>
                                  <TooltipHost
                                    content={nestedArrays.find((obj) => obj?.request === "Publish")?.client}
                                    id="myPersonaTooltip"
                                  >
                                    <Persona
                                      size={PersonaSize.size32}
                                      presence={PersonaPresence.none}
                                      imageUrl={
                                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                                        `${nestedArrays.find((obj) => obj?.request === "Publish")?.ToEmail}`
                                      }
                                    /> </TooltipHost>
                                </> : ''}

                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Distribute") ?
                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Distribute")?.Dev}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Distribute")?.FromEmail}`
                                    }
                                  />

                                </TooltipHost>
                              </>
                              : ''}
                          </td>
                          <td  >



                            {nestedArrays.find((obj) => obj?.request === "Distribute") ?

                              <>
                                <TooltipHost
                                  content={nestedArrays.find((obj) => obj?.request === "Distribute")?.client}
                                  id="myPersonaTooltip"
                                >

                                  <Persona
                                    size={PersonaSize.size32}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                      "/_layouts/15/userphoto.aspx?size=S&username=" +
                                      `${nestedArrays.find((obj) => obj?.request === "Distribute")?.ToEmail}`
                                    }
                                  />

                                </TooltipHost>
                              </>

                              : ''}
                          </td>


                        </>


                      </tr>

                    </>
                  })}





                </>

              }) : <tr>
                <td className="typeData" colSpan={14} width="100%">      <Label style={{ color: "#2392B2", textAlign: "center" }}>No Data Found !!!</Label> </td>



              </tr>
              }


            </table>
            {/* </>} */}
          </div>





          {/* Table-Section Ends */}
        </div>

        <div>
          <Modal isOpen={AdpConfirmationPopup.condition} isBlocking={true}>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                marginTop: "30px",
                width: "450px",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-Start",
                  flexDirection: "column",
                  marginBottom: "10px",
                }}
              >
                <Label className={styles.deletePopupTitle}>Confirmation</Label>
                <Label
                  style={{
                    padding: "5px 20px",
                  }}
                  className={styles.deletePopupDesc}
                >
                  Are you sure want to mark as completed?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setAdpConfirmationPopup({ condition: false, isNew: false });
                  // saveDPData();
                  setAdpLoader("startUpLoader");
                  AdpConfirmationPopup.isNew ? adpAddItem() : null
                }}
                className={styles.apDeletePopupYesBtn}
              >
                Yes
              </button>
              <button
                onClick={(_) => {
                  setAdpConfirmationPopup({ condition: false, isNew: false });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        </div>

        {/* Body-Section Ends */}
      </div>
    </>
  );
};

export default ActivityDeliveryPlan;









































// import * as React from "react";
// import { useState, useEffect } from "react";
// import * as moment from "moment";
// import { Web } from "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import "@pnp/sp/fields";
// import "@pnp/sp/site-users/web";
// import './ActivityPlanStyle2.css'
// import {
//   DetailsList,
//   DetailsListLayoutMode,
//   IDetailsListStyles,
//   SelectionMode,
//   Icon,
//   Label,
//   ILabelStyles,
//   Dropdown,
//   IDropdownStyles,
//   NormalPeoplePicker,
//   Persona,
//   PersonaPresence,
//   PersonaSize,
//   DatePicker,
//   Spinner,
//   PrimaryButton,
//   SearchBox,
//   ISearchBoxStyles,
//   TooltipHost,
//   TooltipOverflowMode,
//   TextField,
//   Checkbox,
//   Modal,
// } from "@fluentui/react";

// import Service from "../components/Services";

// import "../ExternalRef/styleSheets/Styles.css";
// import { FontIcon, IIconProps } from "@fluentui/react/lib/Icon";
// import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
// import styles from "./InnovationHubIntranet.module.scss";
// import CustomLoader from "./CustomLoader";
// import alertify from "alertifyjs";
// import "alertifyjs/build/css/alertify.css";
// import { maxBy } from "lodash";
// import * as Excel from "exceljs/dist/exceljs.min.js";
// import * as FileSaver from "file-saver";
// const saveIcon: IIconProps = { iconName: "Save" };
// const editIcon: IIconProps = { iconName: "Edit" };
// const cancelIcon: IIconProps = { iconName: "Cancel" };

// let DateListFormat = "DD/MM/YYYY";
// let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";
// const inputFormat12 = "YYYY-MM-DDTHH:mm:ss[Z]"; // Input format including time and timezone
// const outputFormat12 = "DD/MM/YYYY"; // Desired output format
// const ActivityDeliveryPlan = (props: any) => {
//   // Variable-Declaration-Section Starts
//   //  const webURL_ = "https://ggsaus.sharepoint.com";
//   //  const WeblistURL_ = "Annual Plan";


//   // const sharepointWeb = Web(webURL_);
//   const sharepointWeb = Web(props.URL);
//   const activityPlan_ID = props.ActivityPlanID;

//   const activityPlanListName = "Activity Plan";
//   const adpListName = "Activity Delivery Plan";
//   const templateListName = "Activity Delivery Plan Template";
//   const activityPBListName = "ActivityProductionBoard";

//   let loggeduseremail: string = props.spcontext.pageContext.user.email;

//   const adpCurrentWeekNumber = moment().isoWeek();
//   const adpCurrentYear = moment().year();

//   const allPeoples = props.peopleList;

//   const adpStatusStyle = mergeStyles({
//     textAlign: "center",
//     borderRadius: "25px",
//   });


//   const adpDrpDwnOptns = {
//     developerOptns: [{ key: "All", text: "All" }],
//     stepsOptns: [{ key: "All", text: "All" }],
//     lessonOptns: [{ key: "All", text: "All" }],
//     statusOptns: [{ key: "All", text: "All" }],
//     weekOptns: [{ key: "All", text: "All" }],
//     yearOptns: [{ key: "All", text: "All" }],
//   };
//   const adpFilterKeys = {
//     developer: "All",
//     step: "All",
//     lesson: "",
//     status: "All",
//     week: "All",
//     year: "All",
//   };
//   // const adpFilterKeys = { developer: "All", step: "All", lesson: "All" };

//   // Variable-Declaration-Section Ends
//   // Styles-Section Starts



//   const adpCommonStatusStyle = mergeStyles({
//     textAlign: "center",
//     borderRadius: 25,
//     fontWeight: "600",
//     padding: 3,
//     width: 100,
//     display: "flex",
//     justifyContent: "center",
//   });

//   const adpbuttonStyle = mergeStyles({
//     textAlign: "center",
//     borderRadius: "2px",
//   });
//   const adpbuttonStyleClass = mergeStyleSets({
//     buttonPrimary: [
//       {
//         color: "White",
//         backgroundColor: "#FAA332",
//         borderRadius: "3px",
//         border: "none",
//         marginRight: "10px",
//         selectors: {
//           ":hover": {
//             backgroundColor: "#FAA332",
//             opacity: 0.9,
//             borderRadius: "3px",
//             border: "none",
//             marginRight: "10px",
//           },
//         },
//       },
//       adpbuttonStyle,
//     ],
//     buttonSecondary: [
//       {
//         color: "White",
//         backgroundColor: "#038387",
//         borderRadius: "3px",
//         border: "none",
//         margin: "0 5px",
//         selectors: {
//           ":hover": {
//             backgroundColor: "#038387",
//             opacity: 0.9,
//           },
//         },
//       },
//       adpbuttonStyle,
//     ],
//   });
//   const adpIconStyleClass = mergeStyleSets({
//     navArrow: [
//       {
//         cursor: "pointer",
//         color: "#2392b2",
//         fontSize: 24,
//         marginTop: "3px",
//         marginRight: 12,
//       },
//     ],
//     navArrowDisabled: [
//       {
//         cursor: "pointer",
//         color: "#ababab",
//         fontSize: 24,
//         marginTop: "3px",
//         marginRight: 12,
//       },
//     ],
//     link: [
//       {
//         fontSize: 17,
//         height: 16,
//         width: 16,
//         color: "#fff",
//         backgroundColor: "#038387",
//         cursor: "pointer",
//         padding: 8,
//         borderRadius: 3,
//         marginLeft: 10,
//         ":hover": {
//           backgroundColor: "#025d60",
//         },
//       },
//     ],
//     linkDisabled: [
//       {
//         fontSize: 18,
//         height: 16,
//         width: 19,
//         color: "#fff",
//         backgroundColor: "#ababab",
//         cursor: "not-allowed",
//         padding: 8,
//         borderRadius: 3,
//         marginLeft: 10,
//       },
//     ],
//     refresh: [
//       {
//         fontSize: 18,
//         height: 16,
//         width: 19,
//         color: "#fff",
//         backgroundColor: "#038387",
//         cursor: "pointer",
//         padding: 8,
//         borderRadius: 3,
//         marginTop: 34,
//         ":hover": {
//           backgroundColor: "#025d60",
//         },
//       },
//     ],
//     save: [
//       {
//         fontSize: "18px",
//         color: "#fff",
//         paddingRight: 10,
//       },
//     ],
//     edit: [
//       {
//         fontSize: "18px",
//         color: "#fff",
//         paddingRight: 10,
//       },
//     ],
//     export: [
//       {
//         color: "black",
//         fontSize: "18px",
//         height: 20,
//         width: 20,
//         cursor: "pointer",
//       },
//     ],
//   });

//   // Styles-Section Ends
//   // States-Declaration Starts
//   const [Open, setOpen] = useState([])
//   const [totalHours, setTotalHours] = useState(0)
//   const [adpReRender, setAdpReRender] = useState(true);
//   const [currentUser, setCurrentUser] = useState({});
//   const [activtyPlanItem, setActivtyPlanItem] = useState([]);
//   const [activityPB, setActivityPB] = useState([]);
//   const [group, setgroup] = useState([]);
//   const [adpMasterData, setAdpMasterData] = useState([]);
//   const [adpData, setAdpData] = useState([]);
//   const [adpDropDownOptions, setAdpDropDownOptions] = useState(adpDrpDwnOptns);
//   const [adpFilters, setAdpFilters] = useState(adpFilterKeys);
//   const [adpActivityResponseData, setAdpActivityResponseData] = useState([]);
//   const [adpEditFlag, setAdpEditFlag] = useState(false);
//   const [newDataFlag, setNewDataFlag] = useState(false);
//   const [adpItemAddFlag, setAdpItemAddFlag] = useState(false);
//   const [adpLoader, setAdpLoader] = useState("noLoader");
//   const [adpLoader2, setAdpLoader2] = useState("noLoader");
//   const [adpAutoSave, setAdpAutoSave] = useState(false);
//   const [finalStep, setFinalStep] = useState([]);
//   const [finalStepConst, setFinalStepConst] = useState([]);
//   const [annualPlanData, setAnnualPlanData] = useState({ ba: '', term: 0 });


//   const [adpSDSort, setAdpSDSort] = useState("");
//   const [adpEDSort, setAdpEDSort] = useState("");

//   const [AdpIsCompleted, setAdpIsCompleted] = useState(false);
//   const [StepsArray, setStepsArray] = useState([])
//   const [reRenderState, setreRenderState] = useState(false)
//   const [StepsArrayCustomized, setStepsArrayyCustomized] = useState(
//     [
//       { name: "Draft", key: "Draft" },
//       { name: "Review", key: "Endorsed" },
//       { name: "Edit", key: "Edited" },
//       { name: "Assemble", key: "Assembled" },
//       { name: "Approved ", key: "Signed Off,Publish ready" },
//       { name: "Distribute", key: "Approved" }
//       // { name: "Distribute", key: "Completed" }
//     ]
//   )
//   const [Tabledata, setTabledata] = useState([])
//   const [diliveryPlanNeedToBeUpdate, setDiliveryPlanNeedToBeUpdate] = useState([])
//   const [AdpConfirmationPopup, setAdpConfirmationPopup] = useState({
//     condition: false,
//     isNew: false,
//   });
//   const formatData = async (records: any) => {
//     console.log(records, "form data")
//     const newArray = [];
//     let currentLesson = records[0]?.LessonID;
//     let currentGroup = [{ ...records[0] }];

//     for (let i = 1; i < records.length; i++) {
//       if (records[i].LessonID === currentLesson) {
//         currentGroup.push({ ...records[i] });

//       } else {
//         newArray.push(currentGroup);
//         currentLesson = records[i]?.LessonID;
//         currentGroup = [{ ...records[i] }];
//       }
//     }

//     newArray.push(currentGroup);
//     const stepsSet = new Set(records.map(obj => obj.Steps));
//     const stepsArray = Array.from(stepsSet);
//     await getReviewLogInfo(activityPlan_ID, records)
//     setTabledata([...newArray])

//     setStepsArray(stepsArray)
//     hoursCal([...newArray])

//   }

//   function sortByModifiedDateAscending(arr: any[]) {
//     // Create a copy of the original array to avoid modifying the original data
//     var sortedArray = arr.slice();
//     // Sort the array by "Modified" date in ascending order
//     sortedArray.sort(function (a: any, b: any) {
//       return moment(a.End).diff(moment(b.End));
//     });
//     return sortedArray;
//   }
//   const getReviewLogInfo = (activityPlan_ID: any, records: any) => {
//     // AH: 5    Developer: { name: 'Charlie Archbold', id: 236, email: 'carchbold@goodtogreatschools.org.au' }
//     // EditorId: 236    End: "14/07/2023"    ID: 36352    IsCompleteNew: false    IsCompleteStatus: false    Lesson: "Lesson 26"    LessonID: 26    MaxPH: ""    MinPH: ""    OrderId: 276    PH: 4    PHError: false    PHWeek: null    Project: "Oz-e-Writing Years F-6 Unit 3
//     // Start: "03/07/2023"    Status: "Scheduled"    Steps: "Draft"    Title: "Draft"    Types: "Curriculum (Writing a lesson)"    dateError:
//     // false

//     let diliveryPlanItems = records;
//     console.log(diliveryPlanItems, "inside getReviewLogInfo")
//     let joinReviewLogList = []
//     sharepointWeb.lists
//       .getByTitle("Review Log")
//       .items.select("*")
//       .filter(`AnnualPlanID eq ${activityPlan_ID}`)
//       .top(5000)
//       .get()
//       .then((items: any) => {

//         let reviewLogItems = [];

//         let count = 0
//         items.map((item) => {


//           if (
//             item.auditRequestType == "Review" || item.auditRequestType == "Initial Edit" || item.auditRequestType == "Assemble" || item.auditRequestType == "Sign-off" || item.auditRequestType == "Publish" || item.auditRequestType == "Distribute") {

//             const diliverySteps = diliveryPlanItems.find(element => element.ID == item.DeliveryPlanID);

//             if (diliverySteps == undefined) {
//               console.log("no availble", item.DeliveryPlanID)
//             } else {
//               reviewLogItems.push({
//                 // FromUserId :item.FromUserId,
//                 request: item.auditRequestType,
//                 response: item.auditLastResponse,
//                 Dev: item.auditFrom,
//                 FromEmail: item.FromEmail,
//                 // ToUserId:item.ToUserId,
//                 client: item.auditTo,
//                 ToEmail: item.ToEmail,

//                 Start: item.Created,
//                 End: item.auditSent,
//                 Created: item.Created,
//                 Modified: item.Modified,
//                 ID: item.ID,

//                 Lesson: diliverySteps.Lesson,
//                 LessonID: diliverySteps.LessonID,
//                 Project: diliverySteps.Project,
//                 StepName: diliverySteps.Steps,
//                 StepID: diliverySteps.ID,
//                 DraftStartDate: diliverySteps.DraftStartDate,
//                 DraftEndDate: diliverySteps.DraftEndDate

//               })
//               count++
//             }
//           }
//         })




//         var sortedData = sortByModifiedDateAscending(reviewLogItems);


//         // //step 4
//         // const groupedLessons = {};
//         // // Iterate over each object in the revLog array
//         // sortedData.forEach(item => {
//         //   const { Lesson } = item;

//         //   if (groupedLessons.hasOwnProperty(Lesson)) {
//         //     // If it exists, push the current item to the existing lesson array
//         //     groupedLessons[Lesson].lessonsData.push(item);
//         //   } else {
//         //     // If it doesn't exist, create a new lesson array with the current item
//         //     groupedLessons[Lesson] = {
//         //       Lesson,
//         //       lessonsData: [item]
//         //     };
//         //   }
//         // });

//         //step 5
//         sortedData.sort(function (a, b) {
//           return a.LessonID - b.LessonID;
//         });

//         console.log(sortedData)

//         //  step 6 
//         const respo = getUniqueLessonsWithSteps(diliveryPlanItems) //adp

//         const revLog = formater(sortedData);

//         console.log(revLog, "sorted log with lesson")
//         console.log(respo, "all lesson with thir draft")
//         var uniqueArr = []
//         var sec_arr = []
//         respo.forEach((item) => {
//           const exist_ = revLog.find((x) => x.StepID === item.StepID);
//           if (exist_) {
//             exist_["lesson"] = [...exist_["lesson"]]

//             console.log([...exist_["lesson"]], "exist_")
//             uniqueArr.push({
//               ...exist_
//             })
//             sec_arr.push({
//               ...exist_
//             })
//             console.log(exist_, 'wxist react')
//           } else {
//             uniqueArr.push({
//               LessonID: item.LessonID,
//               Project: item.Project,
//               l_name: item.l_name,
//               StepName: item.StepName,
//               StepID: item.StepID,
//               DraftStartDate: item.DraftStartDate,
//               DraftEndDate: item.DraftEndDate,
//               Developer: item.Developer,
//               lesson: [...item.responses]
//             })
//             sec_arr.push({
//               StepName: item.StepName,
//               StepID: item.StepID,
//               LessonID: item.LessonID,
//               Project: item.Project,
//               l_name: item.l_name,
//               DraftStartDate: item.DraftStartDate,
//               DraftEndDate: item.DraftEndDate,
//               Developer: item.Developer,
//               lesson: [...item.responses]
//             })
//           }
//         })


//         //write function to just rearrange item accordingly



//         console.log(sec_arr, "before")
//         console.log(uniqueArr, "before")
//         const datauniqueArr = uniqueArr;
//         datauniqueArr.map((res, index) => {

//           const lessonData = datauniqueArr[index].lesson; //main array
//           const responseFormator = revManager(lessonData)
//           datauniqueArr[index].lesson = responseFormator
//         })
//         console.log(datauniqueArr, "after")
//         setFinalStep(datauniqueArr)
//         setFinalStepConst(datauniqueArr)
//         setAdpLoader2("noLoader");
//         console.log("no loader active")








//       })
//       .catch((err) => {
//         console.log(err, "error in review log");
//       });




//   }
//   // function filterLessons(data) {
//   //   var stepsFilter = ["Draft", "Final Draft"];

//   //   // Prepare an empty array to hold unique LessonIDs.
//   //   var uniqueIds = [];

//   //   // Loop over data, to add all unique LessonIDs into uniqueIds array
//   //   for (var i = 0; i < data.length; i++) {
//   //     if (uniqueIds.indexOf(data[i].LessonID) === -1) {
//   //       uniqueIds.push(data[i].LessonID);
//   //     }
//   //   }

//   //   var result = [];

//   //   // Iterate over uniqueIds array
//   //   for (var j = 0; j < uniqueIds.length; j++) {
//   //     // Prepare an array for storing matching responses
//   //     var responses = [];

//   //     // Loop over the data again and add objects with matching LessonID and Steps
//   //     for (var k = 0; k < data.length; k++) {
//   //       if (data[k].LessonID === uniqueIds[j] && stepsFilter.indexOf(data[k].Steps) !== -1) {
//   //         responses.push(data[k]);
//   //       }
//   //     }

//   //     // Add responses array to the result array with LessonID.
//   //     result.push({
//   //       LessonID: uniqueIds[j],
//   //       responses: responses
//   //     });
//   //   }

//   //   return result;
//   // }


//   var objectRef = [
//     {
//       type: "Professional Learning (Lessons)",
//       draft: "Write lesson outline complete"

//     },
//     {
//       type: "Professional Learning (Practice Lessons)",
//       draft: "Write lesson outline complete"
//     },
//     {
//       type: "Event",
//       draft: "Event brief Marketing brief Budget"

//     }
//     ,
//     {
//       type: "Professional Learning (Survey)",
//       draft: "Draft"

//     }, {
//       type: "Curriculum (Survey)",
//       draft: "Draft"
//     },
//     {
//       type: "Marketing (Survey)",
//       draft: "Draft"
//     },
//     {
//       type: "Content Creation (Survey)",
//       draft: "Draft"
//     },
//     {
//       type: "Curriculum (Writing a lesson)",
//       draft: "Draft"
//     },
//     {
//       type: "Marketing (Starting a marketing campaign)",
//       draft: "Produce a product board"
//     },
//     {
//       type: "Marketing (Creating marketing collateral)",
//       draft: "Draft copy"
//     },
//     {
//       type: "Marketing (Delivering marketing collateral)",
//       draft: "Review signed off campaign strategy"
//     },
//     {
//       type: "Marketing (Promoting through the media)",
//       draft: "Build media kit"
//     },
//     {
//       type: "Marketing (Deliver logistics of events)",
//       draft: "Event brief"
//     },



//     {
//       type: "Content Creation (Sourcing digital content)",
//       draft: "Final draft"
//     },
//     {
//       type: "Content Creation (Build a video script)",
//       draft: "Draft script"
//     },
//     {
//       type: "Content Creation (Compile video from footage)",
//       draft: "Additional drafts"
//     },
//     {
//       type: "Content Creation (Shoot video footage)",
//       draft: "Film new content"
//     },
//     {
//       type: "Content Creation (Small graphic)",
//       draft: "Draft"
//     },
//     {
//       type: "Content Creation (Medium graphic)",
//       draft: "Draft"
//     },
//     {
//       type: "Content Creation (Digital Database)",
//       draft: "Final draft"
//     },
//     {
//       type: "Content Creation (Video Content Producer)",
//       draft: "Draft"
//     },



//     {
//       type: "Curriculum (Oz-e-Maths Swap Outs)",
//       draft: "Draft"
//     },
//     {
//       type: "Curriculum (General)",
//       draft: "Draft"
//     },
//     {
//       type: "School Improvement (SCM)",
//       draft: "Draft development"
//     },
//     {
//       type: "Curriculum (Overview)",
//       draft: "Draft"
//     },
//     {
//       type: "School Partnerships",
//       draft: "Analyse data"
//     },
//     {
//       type: "Curriculum Teaching Guide",
//       draft: "Draft"
//     },
//     {
//       type: "Curriculum Student Workbook",
//       draft: "Draft"
//     },
//     {
//       type: "Content Creation (Remote filming)",
//       draft: "Review Video requirements form with Developer"
//     },


//     {
//       type: "Professional Learning (Modules)",
//       draft: "Additional drafts"

//     },
//     {
//       type: "Business Services (Delivery a board meeting)",
//       draft: "CEO to review draft minutes"

//     },
//     {
//       type: "Content Creation (Large graphic)",
//       draft: "Draft"

//     },
//     {
//       type: "School Partnerships V2",
//       draft: "Analyse data and write WDRR"

//     }
//   ]

//   function getString(input: string | string[]): string {
//     // Check if input is an array
//     if (Array.isArray(input)) {
//       // If input is an array, loop through the elements
//       for (let i = 0; i < input.length; i++) {
//         // If the element is a string, return it
//         if (typeof input[i] === 'string') {
//           return input[i];
//         }
//       }
//     }

//     // If input is a string, return it
//     if (typeof input === 'string') {
//       return input;
//     }

//     // If no string is found, return an empty string
//     return '';
//   }

//   function getUniqueLessonsWithSteps(data) {
//     console.log(data, "rerere")
//     let unique_arr = []
//     data.forEach(element => {
//       if (unique_arr.findIndex(x => x.StepID === element.StepID) === -1) {
//         unique_arr.push({
//           LessonID: element.LessonID,
//           StepID: element.ID,
//           StepName: element.Steps,
//           //add here step ids , 
//           //also add here date column for steps 
//           Project: element.Project,
//           l_name: element.Lesson,
//           DraftStartDate: element.DraftStartDate,
//           DraftEndDate: element.DraftEndDate,
//           Developer: element.Developer,
//           responses: []
//         })
//       }
//     });

//     // data.forEach(x => {
//     //   console.log("----------------------------------")
//     //   console.log(x)
//     //   const TypeStringConverted = getString(x.Types)
//     //   console.log(TypeStringConverted)

//     //   const getFruit = objectRef.find(fruit => fruit.type === TypeStringConverted);
//     //   console.log(getFruit)

//     //   if (getFruit && getFruit.draft !== "no" && getFruit.type === TypeStringConverted && getFruit.draft === x.Steps) {

//     //     unique_arr[unique_arr.findIndex(y => y.LessonID === x.LessonID)].responses.push({


//     //       Dev: x?.Developer?.name,
//     //       Dev2: x?.Developer,
//     //       End: x.End, FromEmail: x?.Developer?.email,
//     //       Lesson: x.Lesson, LessonID: x.LessonID, Project: x.Project, Start: x.Start,
//     //       ToEmail: x?.Developer?.email,
//     //       ID: 0,
//     //       ADPId: x.ID,
//     //       request: "Draft",
//     //       response: x.Steps


//     //     })

//     //   }


//     // })

//     console.log(unique_arr, "unique_arr")


//     return unique_arr
//   }




//   const hoursCal = (data) => {
//     data.forEach((lessonData) => {
//       lessonData.forEach((item) => {

//         const hours = calculateWeekdaysWithHours(new Date(
//           moment(item?.Start, DateListFormat).format(DatePickerFormat)
//         ), new Date(
//           moment(item?.End, DateListFormat).format(DatePickerFormat)
//         ));
//         setTotalHours(prevTotalHours => prevTotalHours + hours);
//       });
//     });
//   }
//   window.onbeforeunload = function (e) {
//     if (adpAutoSave) {
//       let dialogText =
//         "You have unsaved changes, are you sure you want to leave?";
//       e.returnValue = dialogText;
//       return dialogText;
//     }
//   };

//   // States-Declaration Ends
//   //Function-Section Starts
//   const generateExcel = () => {



//     const workbook = new Excel.Workbook();
//     const sheet = workbook.addWorksheet('Horizontal Table');
//     // Add Steps row

//     sheet.columns = [
//       { header: "Task", key: "task", width: 25 },
//       { header: "Activity", key: "activity", width: 25 },

//     ];

//     StepsArray.forEach(step => {
//       const cell = sheet.getCell(1, sheet.columnCount + 1);
//       cell.value = step;
//       sheet.mergeCells(1, cell.col, 1, cell.col + 1); // Apply colspan to merged cells
//     });
//     // sheet.addRow(stepsRow);
//     Tabledata.map((data, index) => {


//       const headerRow = [index + 1 == Tabledata.length ? data[0].Types : '', data[0].Lesson];
//       data.forEach(nested => {
//         headerRow.push(nested?.Start, nested?.End);
//       });
//       sheet.addRow(headerRow);
//       const PhotoRow = ['', ''];
//       data.forEach(nested => {
//         PhotoRow.push(nested?.Developer?.name, '');
//       });
//       sheet.addRow(PhotoRow);

//     })
//     workbook.xlsx.writeBuffer().then(buffer => {
//       const file = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
//       FileSaver.saveAs(file, 'horizontal_table.xlsx');
//     });
//   };
//   const calculateWeekdays = (startDate, endDate) => {


//     startDate = convertISODateToStartDate(startDate)
//     endDate = convertISODateToStartDate(endDate)


//     // Validate input
//     if (endDate < startDate)
//       return 0;

//     // Calculate days between dates
//     var millisecondsPerDay = 86400 * 1000; // Day in milliseconds
//     startDate.setHours(0, 0, 0, 1);  // Start just after midnight
//     endDate.setHours(23, 59, 59, 999);  // End just before midnight
//     var diff = endDate - startDate;  // Milliseconds between datetime objects
//     var days = Math.ceil(diff / millisecondsPerDay);

//     // Subtract two weekend days for every week in between
//     var weeks = Math.floor(days / 7);
//     days = days - (weeks * 2);

//     // Handle special cases
//     var startDay = startDate.getDay();
//     var endDay = endDate.getDay();

//     // Remove weekend not previously removed.
//     if (startDay - endDay > 1)
//       days = days - 2;

//     // Remove start day if span starts on Sunday but ends before Saturday
//     if (startDay == 0 && endDay != 6) {
//       days = days - 1;
//     }

//     // Remove end day if span ends on Saturday but starts after Sunday
//     if (endDay == 6 && startDay != 0) {
//       days = days - 1;
//     }

//     return days;
//   }

//   const convertISODateToStartDate = (isoDate) => {
//     // Create a JavaScript Date object from the provided ISO 8601 formatted string
//     let date = new Date(isoDate);

//     // Format the date to the required "DatePickerFormat"
//     let formattedDate = moment(date).format("YYYY-MM-DDT14:00:00Z");

//     // Convert back to JavaScript Date object and return
//     return new Date(formattedDate);
//   };
//   //   const calculateWeekdaysWithHours = (startDate, endDate) => {
//   // // Validate input
//   // if (endDate < startDate)
//   // return 0;

//   // // Calculate hours between dates
//   // var millisecondsPerHour = 60 * 60 * 1000; // Hour in milliseconds

//   // var diff = endDate - startDate;  // Milliseconds between datetime objects
//   // var totalHours = Math.ceil(diff / millisecondsPerHour);

//   // // Calculate start and end day of the week
//   // var startDay = startDate.getDay();
//   // var endDay = endDate.getDay();

//   // // Adjust total hours based on weekends
//   // if (startDay === 0)
//   // startDay = 7; // Sunday is considered as day 7

//   // if (endDay === 0)
//   // endDay = 7; // Sunday is considered as day 7

//   // var weekends = Math.floor((totalHours + startDay - 1) / 24 / 7) * 2;

//   // // Adjust start and end hours
//   // var startHour = startDate.getHours();
//   // var endHour = endDate.getHours();

//   // if (startDay !== 6 && startDay !== 7) {
//   // if (startHour > 0 && startHour < 24) {
//   //   totalHours -= startHour;
//   //   weekends--;
//   // }
//   // }

//   // if (endDay !== 6 && endDay !== 7) {
//   // if (endHour > 0 && endHour < 24) {
//   //   totalHours -= 24 - endHour;
//   //   weekends--;
//   // }
//   // }

//   // // Subtract weekends hours from the total hours
//   // totalHours -= weekends * 24;



//   // return totalHours +"hrs";
//   //   }
//   const calculateWeekdaysWithHours = (startDate, endDate) => {
//     // Validate input
//     if (endDate < startDate)
//       return 0;

//     // Calculate hours and days between dates
//     var millisecondsPerHour = 60 * 60 * 1000; // Hour in milliseconds
//     var millisecondsPerDay = 24 * millisecondsPerHour; // Day in milliseconds

//     startDate.setHours(0, 0, 0, 1);  // Start just after midnight
//     endDate.setHours(23, 59, 59, 999);  // End just before midnight

//     var diff = endDate - startDate;  // Milliseconds between datetime objects
//     var totalHours = Math.ceil(diff / millisecondsPerHour);
//     var totalDays = Math.ceil(diff / millisecondsPerDay);

//     // Subtract two weekend days for every week in between
//     var weeks = Math.floor(totalDays / 7);
//     totalDays = totalDays - (weeks * 2);

//     // Handle special cases
//     var startDay = startDate.getDay();
//     var endDay = endDate.getDay();

//     // Remove weekends not previously removed
//     if (startDay - endDay > 1)
//       totalDays = totalDays - 2;

//     // Remove start day if span starts on Sunday but ends before Saturday
//     if (startDay === 0 && endDay !== 6) {
//       totalDays = totalDays - 1;
//       totalHours = totalHours - (24 - startDate.getHours());
//     }

//     // Remove end day if span ends on Saturday but starts after Sunday
//     if (endDay === 6 && startDay !== 0) {
//       totalDays = totalDays - 1;
//       totalHours = totalHours - endDate.getHours() + 1;
//     }

//     // Adjust hours for complete days
//     totalHours = totalHours - (totalDays * 24);
//     let time = totalDays * 24 + totalHours;

//     return time



//   }



//   const fetchAnnualPlanProduct = async (Project) => {
//     let resp = {
//       ba: '',
//       term: 0
//     }
//     await sharepointWeb.lists.getByTitle("Annual Plan").items.filter(`Title eq '${Project}'`).get().then((items) => {
//       if (items.length > 0) {
//         // Item found
//         const item = items[0];
//         resp = {
//           ba: item.BA_x0020_acronyms,
//           term: item.TermNew.join(', ')
//         }

//         setAnnualPlanData(resp)
//         // Your further code logic here...
//       } else {
//         // Item not found 
//       }
//     }).catch((error) => {
//       console.log("Error occurred:", error);
//     });

//     return resp;
//   }


//   const adpGetCurrentUserDetails = () => {
//     sharepointWeb.currentUser
//       .get()
//       .then((user) => {
//         let adpCurrentUser = {
//           Name: user.Title,
//           Email: user.Email,
//           Id: user.Id,
//         };
//         setCurrentUser({ ...adpCurrentUser });
//       })
//       .catch((err) => {
//         adpErrorFunction(err, "adpGetCurrentUserDetails");
//       });
//   };
//   const getActivityPlanItem = async () => {
//     let _adpItem = [];

//     sharepointWeb.lists
//       .getByTitle(activityPlanListName)
//       .items.getById(activityPlan_ID)
//       .get()
//       .then((item) => {


//         //create function to fetch annual plan product by title fro annul and project from activ plan




//         fetchAnnualPlanProduct(item.Project)



//         _adpItem.push({
//           ID: item.Id ? item.Id : "",
//           Lesson: item.Lessons ? item.Lessons : "",
//           Project: item.Project ? item.Project : "",
//           Product: item.Product ? item.Product : "",
//           ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
//           ProductVersion: item.ProductVersion ? item.ProductVersion : "V1",
//           Types: item.Types ? item.Types : "",
//           Title: item.Title ? item.Title : "",


//           Status: item.Status ? item.Status : null,
//         });

//         let _adpLessons = [];
//         let lessons = _adpItem[0].Lesson.split(";");

//         lessons.forEach((ls) => {
//           _adpLessons.push({
//             ID: ls.split("~")[0],
//             Name: ls.split("~")[1],
//             StartDate: ls.split("~")[2],
//             EndDate: ls.split("~")[3],
//             DeveloperId:
//               ls.split("~")[4] != "NaN" && ls.split("~")[4] != null
//                 ? parseInt(ls.split("~")[4])
//                 : null,
//             DeveloperName:
//               ls.split("~")[4] != "NaN" &&
//                 ls.split("~")[4] != "null" &&
//                 allPeoples.length > 0 &&
//                 allPeoples.filter((ap) => {
//                   return ap.ID == ls.split("~")[4];
//                 }).length > 0
//                 ? allPeoples.filter((ap) => {
//                   return ap.ID == ls.split("~")[4];
//                 })[0].text
//                 : null,
//             DeveloperEmail:
//               ls.split("~")[4] != "NaN" &&
//                 ls.split("~")[4] != "null" &&
//                 allPeoples.length > 0 &&
//                 allPeoples.filter((ap) => {
//                   return ap.ID == ls.split("~")[4];
//                 }).length > 0
//                 ? allPeoples.filter((ap) => {
//                   return ap.ID == ls.split("~")[4];
//                 })[0].secondaryText
//                 : null,
//           });
//         });







//         adpGetData(_adpItem[0], _adpLessons);
//         setActivtyPlanItem([..._adpItem]);
//       })
//       .catch((err) => {
//         adpErrorFunction(err, "getActivityPlanItem");
//       });
//   };
//   const getActivityPBData = () => {
//     sharepointWeb.lists
//       .getByTitle(activityPBListName)
//       .items.filter(
//         `ActivityPlanID eq '${activityPlan_ID}'
//         and Week eq '${adpCurrentWeekNumber}'
//         and Year eq '${adpCurrentYear}'`
//       )
//       .top(5000)
//       .get()
//       .then((items) => {
//         setActivityPB([...items]);
//       })
//       .catch((err) => {
//         adpErrorFunction(err, "getActivityPBData");
//       });
//   };
//   const adpGetData = (adpItem: any, lessons) => {
//     let adpAllitems = [];
//     sharepointWeb.lists
//       .getByTitle(adpListName)
//       .items.select(
//         "*",
//         "Developer/Title",
//         "Developer/Id",
//         "Developer/EMail",
//         "FieldValuesAsText/StartDate",
//         "FieldValuesAsText/EndDate"
//       )
//       .expand("Developer,FieldValuesAsText")
//       .filter(`ActivityPlanID eq ${activityPlan_ID}`)
//       .orderBy("OrderId", true)
//       .top(5000)
//       .get()
//       .then((items) => {
//         if (items.length > 0) {
//           console.log(items, "activity planner")
//           items.forEach((item, index) => {

//             if (Number(item.DeveloperId) == 201) {
//               console.log({
//                 DraftStartDate: item.DraftStartDate
//                   ? moment(item.DraftStartDate, inputFormat12).format(outputFormat12)
//                   : null



//               }, "loewm")
//             }

//             adpAllitems.push({
//               OrderId: index,
//               LessonID: item.LessonID,
//               ID: item.Id ? item.Id : "",
//               Steps: item.Title ? item.Title : "",
//               PH: item.PlannedHours ? item.PlannedHours : "",
//               MinPH: item.MinPH ? item.MinPH : "",
//               EditorId: item.EditorId ? item?.EditorId : "",
//               MaxPH: item.MaxPH ? item.MaxPH : "",
//               Project: item.Project ? item.Project : "",
//               Lesson: item.Lesson ? item.Lesson : "",
//               Types: item.Types ? item.Types : "",
//               Title: item.Title ? item.Title : "",

//               Start: item.StartDate
//                 ? moment(
//                   item["FieldValuesAsText"].StartDate,
//                   DateListFormat
//                 ).format(DateListFormat)
//                 : null,
//               End: item.EndDate
//                 ? moment(
//                   item["FieldValuesAsText"].EndDate,
//                   DateListFormat
//                 ).format(DateListFormat)
//                 : null,
//               Developer: item.DeveloperId
//                 ? {
//                   name: item.Developer.Title,
//                   id: item.Developer.Id,
//                   email: item.Developer.EMail,
//                 }
//                 : "",
//               Status: item.Status ? item.Status : "noData",
//               IsCompleteStatus: item.Status == "Completed" ? true : false,
//               IsCompleteNew: false,
//               AH: item.ActualHours ? item.ActualHours : 0,
//               dateError: false,
//               PHError: false,
//               PHWeek: item.PHWeek ? item.PHWeek : null,
//               DraftStartDate: item.DraftStartDate
//                 ? moment(item.DraftStartDate, inputFormat12).format(outputFormat12)
//                 : null,
//               DraftEndDate:item.DraftEndDate
//               ? moment(item.DraftEndDate, inputFormat12).format(outputFormat12)
//               : null,
//             });
//           });
//           adpGetTemplateData(adpItem, lessons, adpAllitems, items.length);
//           // groups(adpAllitems);
//           // adpGetAllOptions(adpAllitems);

//           // setAdpActivityResponseData([...adpAllitems]);
//           // setAdpData([...adpAllitems]);
//           // setAdpMasterData([...adpAllitems]);
//           // setAdpLoader("noLoader");
//         } else {
//           setNewDataFlag(true);
//           sharepointWeb.lists
//             .getByTitle(templateListName)
//             .items.filter(`Types eq '${adpItem.Types}'`)
//             .orderBy("ID", true)
//             .top(5000)
//             .get()
//             .then((items) => {
//               let count = 0;
//               lessons.forEach((ls) => {
//                 items.forEach((item, index) => {
//                   let PHErrorFlag =
//                     item.MinHours && item.MaxHours
//                       ? adpPHValidationFunction(
//                         parseFloat(item.Hours ? item.Hours : 0),
//                         item.MinHours,
//                         item.MaxHours
//                       )
//                       : false;
//                   // let datediff =
//                   //   new Date(ls.EndDate).getDate() -
//                   //   new Date(ls.StartDate).getDate() +
//                   //   1;
//                   // let Hours = item.Week ? datediff / 7 : item.Hours;
//                   const date1: any = new Date(ls.StartDate);
//                   const date2: any = new Date(ls.EndDate);
//                   const diffTime = Math.abs(date2 - date1);
//                   const diffDays =
//                     Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
//                   const Hours = item.Week ? diffDays / 7 : item.Hours;

//                   adpAllitems.push({
//                     OrderId: count++,
//                     ID: count,
//                     Steps: item.Title ? item.Title : "",
//                     Types: item.Types ? item.Types : "",
//                     Title: item.Title ? item.Title : "",

//                     PH: Hours ? Hours : "",
//                     MinPH: item.MinHours ? item.MinHours : "",
//                     MaxPH: item.MaxHours ? item.MaxHours : "",
//                     Project: adpItem.Project ? adpItem.Project : "",
//                     LessonID: ls.ID,
//                     Lesson: ls.Name ? ls.Name : "",
//                     Start: moment(ls.StartDate).format(DateListFormat),
//                     End: moment(ls.EndDate).format(DateListFormat),
//                     Developer: {
//                       name: ls.DeveloperName,
//                       id: ls.DeveloperId,
//                       email: ls.DeveloperEmail,
//                     },
//                     Status: "Scheduled",
//                     IsCompleteStatus: false,
//                     IsCompleteNew: false,
//                     AH: 0,
//                     dateError: false,
//                     PHError: PHErrorFlag,
//                     PHWeek: item.Week ? item.Week : null,
//                   });
//                 });
//               });
//               groups(adpAllitems);
//               adpGetAllOptions(adpAllitems);

//               // setAdpActivityResponseData([...adpAllitems]);
//               setAdpData([...adpAllitems]);
//               setAdpMasterData([...adpAllitems]);
//               setAdpLoader("noLoader");
//             })
//             .catch((err) => {
//               adpErrorFunction(err, "adpGetData-getTemplateData");
//             });
//         }
//       })
//       .catch((err) => {
//         adpErrorFunction(err, "adpGetData-getADPData");
//       });
//   };

//   //!Update template in the database
//   const adpGetTemplateData = (
//     adpItem: any,
//     lessons,
//     adplistItems: any[],
//     countLists
//   ) => {
//     let adpAllitems = adplistItems;
//     sharepointWeb.lists
//       .getByTitle(templateListName)
//       .items.filter(`Types eq '${adpItem.Types}'`)
//       .orderBy("ID", true)
//       .top(5000)
//       .get()
//       .then((items) => {
//         let count = countLists;
//         lessons.forEach((ls) => {
//           let curLessondata = adplistItems.filter((arr) => {
//             return arr.LessonID == ls.ID;
//           });

//           curLessondata.length == 0 &&
//             items.forEach((item, index) => {
//               let PHErrorFlag =
//                 item.MinHours && item.MaxHours
//                   ? adpPHValidationFunction(
//                     parseFloat(item.Hours ? item.Hours : 0),
//                     item.MinHours,
//                     item.MaxHours
//                   )
//                   : false;
//               const date1: any = new Date(ls.StartDate);
//               const date2: any = new Date(ls.EndDate);
//               const diffTime = Math.abs(date2 - date1);
//               const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
//               const Hours = item.Week ? diffDays / 7 : item.Hours;

//               adpAllitems.push({
//                 OrderId: count++,
//                 ID: 0,
//                 Steps: item.Title ? item.Title : "",
//                 PH: Hours ? Hours : "",
//                 MinPH: item.MinHours ? item.MinHours : "",
//                 MaxPH: item.MaxHours ? item.MaxHours : "",
//                 Project: adpItem.Project ? adpItem.Project : "",
//                 LessonID: ls.ID,
//                 Lesson: ls.Name ? ls.Name : "",
//                 Types: item.Types ? item.Types : "",
//                 Title: item.Title ? item.Title : "",

//                 Start: moment(ls.StartDate).format(DateListFormat),
//                 End: moment(ls.EndDate).format(DateListFormat),
//                 Developer: {
//                   name: ls.DeveloperName,
//                   id: ls.DeveloperId,
//                   email: ls.DeveloperEmail,
//                 },
//                 Status: "Scheduled",
//                 IsCompleteStatus: false,
//                 IsCompleteNew: false,
//                 AH: 0,
//                 dateError: false,
//                 PHError: PHErrorFlag,
//                 PHWeek: item.Week ? item.Week : null,
//               });
//             });
//         });
//         groups(adpAllitems);
//         adpGetAllOptions(adpAllitems);

//         // setAdpActivityResponseData([...adpAllitems]);
//         setAdpData([...adpAllitems]);
//         setAdpMasterData([...adpAllitems]);
//         setAdpLoader("noLoader");
//       })
//       .catch((err) => {
//         adpErrorFunction(err, "adpGetData-getTemplateData");
//       });
//   };

//   const adpGetAllOptions = (allItems: any) => {
//     allItems.forEach((item: any) => {
//       if (
//         adpDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
//           return developerOptn.key == item.Developer.name;
//         }) == -1 &&
//         item.Developer.name
//       ) {
//         adpDrpDwnOptns.developerOptns.push({
//           key: item.Developer.name,
//           text: item.Developer.name,
//         });
//       }

//       if (
//         adpDrpDwnOptns.stepsOptns.findIndex((stepsOptn) => {
//           return stepsOptn.key == item.Steps;
//         }) == -1 &&
//         item.Steps
//       ) {
//         adpDrpDwnOptns.stepsOptns.push({
//           key: item.Steps,
//           text: item.Steps,
//         });
//       }
//       if (
//         adpDrpDwnOptns.statusOptns.findIndex((statsOptn) => {
//           return statsOptn.key == item.Status;
//         }) == -1 &&
//         item.Status
//       ) {
//         adpDrpDwnOptns.statusOptns.push({
//           key: item.Status,
//           text: item.Status,
//         });
//       }

//       if (
//         adpDrpDwnOptns.lessonOptns.findIndex((lessonOptn) => {
//           return lessonOptn.key == item.Lesson;
//         }) == -1 &&
//         item.Lesson
//       ) {
//         adpDrpDwnOptns.lessonOptns.push({
//           key: item.Lesson,
//           text: item.Lesson,
//         });
//       }
//     });

//     let maxWeek =
//       parseInt(adpFilters.year) == moment().year() ? moment().isoWeek() : 53;

//     for (var i = 1; i <= maxWeek; i++) {
//       adpDrpDwnOptns.weekOptns.push({
//         key: i.toString(),
//         text: i.toString(),
//       });
//     }
//     for (var i = 2020; i <= moment().year(); i++) {
//       adpDrpDwnOptns.yearOptns.push({
//         key: i.toString(),
//         text: i.toString(),
//       });
//     }

//     let unsortedFilterKeys = adpSortingFilterKeys(adpDrpDwnOptns);
//     setAdpDropDownOptions({ ...unsortedFilterKeys });
//   };
//   const adpSortingFilterKeys = (unsortedFilterKeys: any) => {
//     const sortFilterKeys = (a, b) => {
//       if (a.text < b.text) {
//         return -1;
//       }
//       if (a.text > b.text) {
//         return 1;
//       }
//       return 0;
//     };

//     if (
//       unsortedFilterKeys.developerOptns.some((managerOptn) => {
//         return (
//           managerOptn.text.toLowerCase() ==
//           props.spcontext.pageContext.user.displayName.toLowerCase()
//         );
//       })
//     ) {
//       unsortedFilterKeys.developerOptns.shift();
//       let loginUserIndex = unsortedFilterKeys.developerOptns.findIndex(
//         (user) => {
//           return (
//             user.text.toLowerCase() ==
//             props.spcontext.pageContext.user.displayName.toLowerCase()
//           );
//         }
//       );
//       let loginUserData = unsortedFilterKeys.developerOptns.splice(
//         loginUserIndex,
//         1
//       );

//       unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
//       unsortedFilterKeys.developerOptns.unshift(loginUserData[0]);
//       unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
//     } else {
//       unsortedFilterKeys.developerOptns.shift();
//       unsortedFilterKeys.developerOptns.sort(sortFilterKeys);
//       unsortedFilterKeys.developerOptns.unshift({ key: "All", text: "All" });
//     }

//     unsortedFilterKeys.statusOptns.shift();
//     unsortedFilterKeys.statusOptns.sort(sortFilterKeys);
//     unsortedFilterKeys.statusOptns.unshift({ key: "All", text: "All" });

//     unsortedFilterKeys.stepsOptns.shift();
//     unsortedFilterKeys.stepsOptns.sort(sortFilterKeys);
//     unsortedFilterKeys.stepsOptns.unshift({ key: "All", text: "All" });

//     unsortedFilterKeys.lessonOptns.shift();
//     unsortedFilterKeys.lessonOptns.sort(sortFilterKeys);
//     unsortedFilterKeys.lessonOptns.unshift({ key: "All", text: "All" });

//     return unsortedFilterKeys;
//   };
//   const adpListFilter = (key: string, option: any) => {
//     let arrBeforeFilter = [...adpData];

//     let tempFilterKeys = { ...adpFilters };
//     tempFilterKeys[key] = option;

//     if (tempFilterKeys.developer != "All") {
//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         return arr.Developer.name == tempFilterKeys.developer;
//       });
//     }

//     if (tempFilterKeys.step != "All") {
//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         return arr.Steps == tempFilterKeys.step;
//       });
//     }
//     if (tempFilterKeys.status != "All") {
//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         return arr.Status == tempFilterKeys.status;
//       });
//     }
//     if (tempFilterKeys.lesson) {
//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         return arr.Lesson.toLowerCase().includes(
//           tempFilterKeys.lesson.toLowerCase()
//         );
//       });
//     }

//     if (tempFilterKeys.week != "All") {
//       let year =
//         tempFilterKeys.year == "All" ? moment().year() : tempFilterKeys.year;

//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         let start = moment(arr.Start, DateListFormat)
//           .year()
//           .toString()
//           .concat(
//             (
//               "0" + moment(arr.Start, DateListFormat).isoWeek().toString()
//             ).slice(-2)
//           );
//         let end = moment(arr.End, DateListFormat)
//           .year()
//           .toString()
//           .concat(
//             ("0" + moment(arr.End, DateListFormat).isoWeek().toString()).slice(
//               -2
//             )
//           );
//         let today = year
//           .toString()
//           .concat(("0" + tempFilterKeys.week.toString()).slice(-2));

//         return (
//           parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
//         );
//       });
//     }

//     if (tempFilterKeys.year != "All") {
//       arrBeforeFilter = arrBeforeFilter.filter((arr) => {
//         let start = moment(arr.Start, DateListFormat).year().toString();

//         let end = moment(arr.End, DateListFormat).year().toString();

//         let today = tempFilterKeys.year.toString();

//         return (
//           parseInt(today) >= parseInt(start) && parseInt(today) <= parseInt(end)
//         );
//       });
//     }

//     groups([...arrBeforeFilter]);
//     // setAdpActivityResponseData([...arrBeforeFilter]);
//     setAdpFilters({ ...tempFilterKeys });
//   };
//   const overallPlannedHours = () => {
//     let ph = 0;
//     if (adpData.length > 0) {
//       adpData.forEach((data) => {
//         ph += data.PH ? data.PH : 0;
//       });
//     }
//     return ph;
//   };
//   const overallActualHours = () => {
//     let ah = 0;
//     if (adpData.length > 0) {
//       adpData.forEach((data) => {
//         ah += data.AH ? data.AH : 0;
//       });
//     }
//     return ah;
//   };

//   const updateDeveloperData = () => {

//   }


//   const adpActivityResponseHandler = (id: number, key: string, value: any) => {
//     //////.......................................... 


//     let tempDeveloper = [];
//     console.log(adpData, "adpActivityResponse")
//     let Index = adpData.findIndex((data) => data.OrderId == id);
//     let disIndex = adpActivityResponseData.findIndex(
//       (data) => data.OrderId == id
//     );

//     let adpBeforeData = adpData[Index];

//     if (key == "Developer") {
//       if (value) {
//         tempDeveloper = allPeoples.filter((people) => {
//           return people.ID == value;
//         });
//       }
//     }

//     let dateErrorFlag = adpDateValidationFunction(
//       key == "Start"
//         ? moment(value).format("YYYY/MM/DD")
//         : moment(adpBeforeData.Start, DateListFormat).format("YYYY/MM/DD"),
//       key == "End"
//         ? moment(value).format("YYYY/MM/DD")
//         : moment(adpBeforeData.End, DateListFormat).format("YYYY/MM/DD")
//     );
//     let PHErrorFlag =
//       key == "PH"
//         ? adpPHValidationFunction(
//           parseFloat(value),
//           adpBeforeData.MinPH,
//           adpBeforeData.MaxPH
//         )
//         : adpBeforeData.PHError;

//     const date1: any = new Date(key == "Start" ? value : adpBeforeData.Start);
//     const date2: any = new Date(key == "End" ? value : adpBeforeData.End);
//     const diffTime = Math.abs(date2 - date1);
//     const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
//     const Hours = adpBeforeData.PHWeek ? diffDays / 7 : adpBeforeData.PH;

//     let adpOnchangeData = {
//       LessonID: adpBeforeData.LessonID,
//       OrderId: adpBeforeData.OrderId,
//       ID: adpBeforeData.ID,
//       Steps: adpBeforeData.Steps,
//       PH: key == "PH" ? value : Hours,
//       MinPH: adpBeforeData.MinPH,
//       MaxPH: adpBeforeData.MaxPH,
//       Project: adpBeforeData.Project,
//       Lesson: adpBeforeData.Lesson,
//       Start:
//         key == "Start"
//           ? moment(value).format(DateListFormat)
//           : adpBeforeData.Start,
//       End:
//         key == "End" ? moment(value).format(DateListFormat) : adpBeforeData.End,
//       Developer:
//         key == "Developer"
//           ? {
//             name: tempDeveloper.length > 0 ? tempDeveloper[0].text : "",
//             id: tempDeveloper.length > 0 ? tempDeveloper[0].ID : null,
//             email:
//               tempDeveloper.length > 0 ? tempDeveloper[0].secondaryText : "",
//           }
//           : adpBeforeData.Developer,
//       Status: adpBeforeData.Status,
//       IsCompleteStatus:
//         key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteStatus,
//       IsCompleteNew:
//         key == "IsCompleteStatus" ? value : adpBeforeData.IsCompleteNew,
//       dateError: dateErrorFlag,
//       PHError: PHErrorFlag,
//       PHWeek: adpBeforeData.PHWeek,
//     };

//     adpData[Index] = adpOnchangeData;
//     adpActivityResponseData[disIndex] = adpOnchangeData;

//     let isCompleteDetails = adpData.filter((arr) => {
//       return arr.IsCompleteStatus == true;
//     });
//     adpData.length == isCompleteDetails.length
//       ? setAdpIsCompleted(true)
//       : setAdpIsCompleted(false);

//     setAdpData([...adpData]);
//     groups([...adpActivityResponseData]);
//     // setAdpActivityResponseData([...adpActivityResponseData]);
//   };
//   const adpAddItem = () => {
//     let successCount = 0;

//     let completionDetails = adpData.filter((arr) => {
//       return arr.IsCompleteStatus == true;
//     });

//     let completionValue =
//       adpData.length > 0 && completionDetails.length > 0
//         ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
//         : 0;

//     adpData.forEach(async (response: any, index: number) => {
//       let strDSNA: string = `${response.Developer.id ? response.Developer.id : null
//         }-0`;

//       let statusValue = response.IsCompleteStatus
//         ? "Completed"
//         : response.Status
//           ? response.Status
//           : null;

//       let responseData = {
//         ActivityPlanID: activityPlan_ID ? activityPlan_ID.toString() : "",
//         Title: response.Steps ? response.Steps : "",
//         PlannedHours: response.PH ? response.PH : 0,
//         MinPH: response.MinPH ? response.MinPH : 0,
//         MaxPH: response.MaxPH ? response.MaxPH : 0,
//         ProjectVersion: activtyPlanItem[0].ProjectVersion
//           ? activtyPlanItem[0].ProjectVersion
//           : "V1",
//         ProductVersion: activtyPlanItem[0].ProductVersion
//           ? activtyPlanItem[0].ProductVersion
//           : "V1",
//         Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
//         Types: activtyPlanItem[0].Types ? activtyPlanItem[0].Types : "",



//         Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
//         Lesson: response.Lesson ? response.Lesson : "",
//         StartDate: response.Start
//           ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
//           : moment().format("YYYY-MM-DD"),
//         EndDate: response.End
//           ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
//           : moment().format("YYYY-MM-DD"),
//         DeveloperId: response.Developer.id ? response.Developer.id : null,
//         // Status: "Scheduled",
//         Status: statusValue,
//         ActualHours: 0,
//         OrderId: response.OrderId,
//         LessonID: response.LessonID ? response.LessonID : null,
//         PHWeek: response.PHWeek ? response.PHWeek : null,
//         SPFxFilter: strDSNA,
//       };

//       // debugger;

//       await sharepointWeb.lists
//         .getByTitle(adpListName)
//         .items.add(responseData)
//         .then((item) => {
//           successCount++;
//           adpData[index].ID = item.data.Id;
//           adpData[index].Status = statusValue;
//           adpData[index].IsCompleteNew = false;

//           if (adpData.length == successCount) {
//             let apCompletedValue = adpData.filter((arr) => {
//               return arr.IsCompleteStatus == true;
//             });
//             if (
//               adpData.length == apCompletedValue.length &&
//               activtyPlanItem[0].Status != "Completed"
//             ) {
//               sharepointWeb.lists
//                 .getByTitle(activityPlanListName)
//                 .items.getById(activityPlan_ID)
//                 .update({
//                   Status: "Completed",
//                   Completion: 100,
//                   CompletedDate: moment().format("YYYY-MM-DD"),
//                 })
//                 .then((e) => { })
//                 .catch((err) => {
//                   adpErrorFunction(err, "saveDPData-getAPItem");
//                 });
//             } else {
//               sharepointWeb.lists
//                 .getByTitle(activityPlanListName)
//                 .items.getById(activityPlan_ID)
//                 .update({
//                   Completion: completionValue,
//                 })
//                 .then((e) => { })
//                 .catch((err) => {
//                   adpErrorFunction(err, "saveDPData-getAPItem");
//                 });
//             }

//             const newData = _copyAndSort(adpData, "OrderId", false);

//             adpGetAllOptions([...newData]);
//             setAdpMasterData([...newData]);
//             setAdpData([...newData]);
//             groups([...newData]);

//             // setAdpActivityResponseData([...adpData]);
//             setNewDataFlag(false);
//             setAdpItemAddFlag(true);
//             setAdpEditFlag(false);
//             setAdpLoader("noLoader");
//             AddSuccessPopup();
//           }
//         })
//         .catch((err) => {
//           adpErrorFunction(err, "adpAddItem");
//         });
//     });
//   };

//   const adpUpdateItem_Old = () => {
//     let responseDataArr = [];
//     let newArr = [...adpData];
//     let successCount = 0;

//     let selected = [];

//     adpData.forEach((response: any, index: number) => {
//       let targetStatus = newArr.filter((arr) => {
//         return arr.ID == response.ID;
//       });

//       let strDSNA: string = `${response.Developer.id}-${targetStatus[0].Status == "Completed" ? 1 : 0
//         }`;

//       let responseData = {
//         ProjectVersion: activtyPlanItem[0].ProjectVersion
//           ? activtyPlanItem[0].ProjectVersion
//           : "V1",
//         ProductVersion: activtyPlanItem[0].ProductVersion
//           ? activtyPlanItem[0].ProductVersion
//           : "V1",
//         Product: activtyPlanItem[0].Product ? activtyPlanItem[0].Product : "",
//         Project: activtyPlanItem[0].Project ? activtyPlanItem[0].Project : "",
//         PlannedHours: response.PH ? response.PH : 0,
//         StartDate: response.Start
//           ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
//           : null,
//         EndDate: response.End
//           ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
//           : null,
//         DeveloperId: response.Developer.id ? response.Developer.id : null,
//         SPFxFilter: strDSNA,
//       };

//       responseDataArr.push(responseData);

//       sharepointWeb.lists
//         .getByTitle(adpListName)
//         .items.getById(response.ID)
//         .update(responseData)
//         .then(() => {
//           successCount++;
//           let newDeveloperDetails = {};

//           let targetIndex = newArr.findIndex((arr) => arr.ID == response.ID);
//           let targetItem = newArr.filter((arr) => {
//             return arr.ID == response.ID;
//           });

//           if (response.Developer.id) {
//             let newDeveloper = allPeoples.filter((people) => {
//               return people.ID == response.Developer.id;
//             });
//             newDeveloperDetails = {
//               name: newDeveloper[0].text,
//               id: newDeveloper[0].ID,
//               email: newDeveloper[0].secondaryText,
//             };
//           } else {
//             newDeveloperDetails = {
//               name: null,
//               id: null,
//               email: null,
//             };
//           }

//           newArr[targetIndex] = {
//             OrderId: response.OrderId,
//             ID: targetItem[0].ID ? targetItem[0].ID : "",
//             Steps: targetItem[0].Steps ? targetItem[0].Steps : "",
//             PH: response.PH ? response.PH : "",
//             MinPH: targetItem[0].MinPH ? targetItem[0].MinPH : "",
//             MaxPH: targetItem[0].MaxPH ? targetItem[0].MaxPH : "",
//             Project: targetItem[0].Project ? targetItem[0].Project : "",
//             LessonID: targetItem[0].LessonID ? targetItem[0].LessonID : null,
//             Lesson: targetItem[0].Lesson ? targetItem[0].Lesson : "",
//             Start: response.Start ? response.Start : targetItem[0].Start,
//             End: response.End ? response.End : targetItem[0].End,
//             Developer: newDeveloperDetails,
//             Status: targetItem[0].Status ? targetItem[0].Status : "",
//             AH: targetItem[0].AH ? targetItem[0].AH : "",
//             dateError: false,
//             PHError: false,
//             PHWeek: targetItem[0].PHWeek ? targetItem[0].PHWeek : null,
//           };

//           let filteredPB = activityPB.filter((pb) => {
//             return pb.ActivityDeliveryPlanID == newArr[targetIndex].ID;
//           });

//           selected.push([...filteredPB]);

//           if (filteredPB.length > 0) {
//             sharepointWeb.lists
//               .getByTitle(activityPBListName)
//               .items.getById(filteredPB[0].ID)
//               .update({
//                 PlannedHours: response.PH ? response.PH : 0,
//                 StartDate: response.Start
//                   ? moment(response.Start, DateListFormat).format("YYYY-MM-DD")
//                   : null,
//                 EndDate: response.End
//                   ? moment(response.End, DateListFormat).format("YYYY-MM-DD")
//                   : null,
//                 DeveloperId: response.Developer.id
//                   ? response.Developer.id
//                   : null,
//               })
//               .then((e) => { })
//               .catch((err) => {
//                 adpErrorFunction(err, "adpUpdateItem-updateAPBList");
//               });
//           }

//           if (adpActivityResponseData.length == successCount) {
//             adpGetAllOptions(newArr);
//             setAdpEditFlag(false);
//             setAdpMasterData([...newArr]);
//             setAdpLoader("noLoader");
//             AddSuccessPopup();
//           }
//         })
//         .catch((err) => {
//           adpErrorFunction(err, "adpUpdateItem-updateATPList");
//         });
//     });
//   };
//   const show = (position) => {
//     let data = Open.filter((i) => i === position);

//     if (data.length > 0) {
//       setOpen(Open.filter((i) => i != position));
//     } else {
//       setOpen([...Open, position]);
//     }

//   };

//   function convertDateFormat(inputDate) {
//     const formattedDate = moment(inputDate).utc().format('YYYY-MM-DDTHH:mm:ss[Z]');
//     return formattedDate;
//   }

//   function convertToCustomFormat(inputDate, outputFormat) {
//     const formattedDate = moment(inputDate).format(outputFormat);
//     return formattedDate;
//   }
//   const adpUpdateItem = () => {
//     // let responseDataArr = [];
//     // let newArr = [...adpData];
//     // let successCount = 0;

//     // let completionDetails = adpData.filter((arr) => {
//     //   return arr.IsCompleteStatus == true;
//     // });

//     // let completionValue =
//     //   adpData.length > 0 && completionDetails.length > 0
//     //     ? ((completionDetails.length / adpData.length) * 100).toFixed(2)
//     //     : 0;
//     // console.log(diliveryPlanNeedToBeUpdate)
//     console.log(diliveryPlanNeedToBeUpdate), "1234567890"
//     diliveryPlanNeedToBeUpdate.forEach((item, index) => {
//       if (item.StepID) {

//         let inputDate1
//         if (item.DraftStartDate) {
//           inputDate1 = item.DraftStartDate
//         } else {
//           inputDate1 = ''
//         }

//         const inputFormat1 = "DD/MM/YYYY";
//         const outputFormat1 = "YYYY-MM-DDTHH:mm:ss[Z]";

//         const convertedDate1 = convertDateFormat2(inputDate1, inputFormat1, outputFormat1);

//         let inputDate2;
//         if (item.DraftEndDate) {
//           inputDate2 = item.DraftEndDate
//         } else {
//           inputDate2 = ''
//         }




//         const inputFormat2 = "DD/MM/YYYY";
//         const outputFormat2 = "YYYY-MM-DDTHH:mm:ss[Z]";

//         const convertedDate2 = convertDateFormat2(inputDate2, inputFormat2, outputFormat2);




//         let responseData2 = {
//           // ProjectVersion: "v2",
//           // ProductVersion: "v2",
//           // Product: "Mastery Teaching Pathway",
//           // Project: "Coach Positive High Expectations ",

//           // PlannedHours: 0,
//           DraftStartDate: convertedDate1,
//           DraftEndDate: convertedDate2,
//           DeveloperId: item?.Developer?.id,


//         }

//         console.log(responseData2)
//         sharepointWeb.lists
//           .getByTitle(adpListName)
//           .items.getById(item.StepID)
//           .update(responseData2)
//           .then(() => {
//             console.log("success");


//           })
//           .catch((err) => {
//             adpErrorFunction(err, "adpUpdateItem-updateATPList");
//           });
//       }
//     })

//     setFinalStepConst(finalStep)
//     setAdpEditFlag(false);
//     setAdpAutoSave(false);

//   };









//   const adpDateValidationFunction = (startDate: any, EndDate: any) => {
//     if (startDate != null && EndDate != null) {
//       if (startDate > EndDate) {
//         return true;
//       } else {
//         return false;
//       }
//     } else {
//       return false;
//     }
//   };
//   const adpPHValidationFunction = (val, min, max) => {
//     if (val >= min && val <= max) {
//       return false;
//     } else {
//       return true;
//     }
//   };


//   function convertDateFormat2(inputDate, inputFormat, outputFormat) {
//     const parsedDate = moment(inputDate, inputFormat);
//     const formattedDate = parsedDate.utc().format(outputFormat);
//     return formattedDate;
//   }



//   const dateFormater = (date: Date): string => {
//     return date ? moment(date).format("DD/MM/YYYY") : "";
//   };
//   const GetUserDetails = (filterText, currentPersonas) => {
//     let _allPeoples = allPeoples;

//     _allPeoples = _allPeoples.filter((arr) => {
//       return arr.text.toLowerCase().indexOf("archive") == -1;
//     });

//     if (currentPersonas.length > 0) {
//       _allPeoples = _allPeoples.filter(
//         (arr) => !currentPersonas.some((persona) => persona.ID == arr.ID)
//       );
//     }

//     var result = _allPeoples.filter(
//       (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
//     );

//     return result.filter((item) =>
//       doesTextStartWith(item.text as string, filterText)
//     );
//   };
//   const doesTextStartWith = (text: string, filterText: string) => {
//     return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
//   };
//   const adpErrorFunction = (error: any, functionName: string) => {
//     console.log(error);

//     let response = {
//       ComponentName: "Activity delivery plan",
//       FunctionName: functionName,
//       ErrorMessage: JSON.stringify(error["message"]),
//       Title: loggeduseremail,
//     };

//     Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
//       () => {
//         setAdpLoader("noLoader");
//         setAdpEditFlag(false);
//         ErrorPopup();
//         setAdpReRender(!adpReRender);
//       }
//     );
//   };
//   const AddSuccessPopup = () => (
//     alertify.set("notifier", "position", "top-right"),
//     alertify.success("Activity planner is successfully submitted !!!")
//   );
//   const ErrorPopup = () => (
//     alertify.set("notifier", "position", "top-right"),
//     alertify.error("Something when error, please contact system admin.")
//   );

//   const sortingFunction = (columnName, sortType): void => {
//     let tempArr = adpData;
//     let tempDisArr = adpActivityResponseData;

//     const newDisData = _copyAndSort(
//       tempDisArr,
//       columnName,
//       sortType == "desc" ? true : false
//     );
//     const newData = _copyAndSort(
//       tempArr,
//       columnName,
//       sortType == "desc" ? true : false
//     );

//     setAdpData([...newData]);
//     groups([...newDisData]);
//   };

//   function _copyAndSort<T>(
//     items: T[],
//     columnKey: string,
//     isSortedDescending?: boolean
//   ): T[] {
//     let key = columnKey as keyof T;
//     return items
//       .slice(0)
//       .sort((a: T, b: T) =>
//         (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
//       );
//   }

//   const groups = (records) => {
//     console.log(records, "ITS HRERE")
//     let reOrderedRecords = [];

//     let Uniquelessons = records.reduce(function (item, e1) {
//       var matches = item.filter(function (e2) {
//         return e1.Lesson === e2.Lesson;
//       });

//       if (matches.length == 0) {
//         item.push(e1);
//       }
//       return item;
//     }, []);

//     Uniquelessons.forEach((ul) => {
//       let curLesson = records.filter((arr) => {
//         return arr.Lesson == ul.Lesson;
//       });
//       reOrderedRecords = reOrderedRecords.concat(curLesson);
//     });
//     groupsforDL(reOrderedRecords);
//   };

//   const groupsforDL = (records) => {
//     let newRecords = [];
//     records.forEach((rd, index) => {
//       newRecords.push({
//         Lesson: rd.Lesson,
//         indexValue: index,
//       });
//     });

//     let varGroup = [];
//     let Uniquelessons = newRecords.reduce(function (item, e1) {
//       var matches = item.filter(function (e2) {
//         return e1.Lesson === e2.Lesson;
//       });

//       if (matches.length == 0) {
//         item.push(e1);
//       }
//       return item;
//     }, []);

//     Uniquelessons.forEach((ul) => {
//       let lessonLength = newRecords.filter((arr) => {
//         return arr.Lesson == ul.Lesson;
//       }).length;
//       varGroup.push({
//         key: ul.Lesson,
//         name: ul.Lesson,
//         startIndex: ul.indexValue,
//         count: lessonLength,
//       });
//     });
//     setAdpActivityResponseData([...records]);
//     let filterRec = records.filter(item => {
//       return item.Steps == 'Draft' ||
//         item.Steps == 'Review' ||
//         item.Steps == 'Edit' ||
//         item.Steps == 'Assemble' ||
//         item.Steps == 'Sign Off' ||
//         item.Steps == 'Publish' ||
//         item.Steps == 'Distribute';
//     });
//     formatData(records)
//     setgroup([...varGroup]);
//   };

//   function checkIfLessonExists(StepID: any, endResp: any) {
//     for (let i = 0; i < endResp.length; i++) {
//       if (endResp[i].StepID === StepID) {
//         return true;
//       }
//     }
//     return false;
//   }

//   const formater = (arrayData: any) => {
//     const endResp: any = [];
//     arrayData.map((x: any) => {
//       if (!checkIfLessonExists(x.StepID, endResp)) {
//         endResp.push({
//           LessonID: x.LessonID,
//           l_name: x.Lesson,
//           Project: x.Project,
//           StepID: x.StepID,
//           StepName: x.StepName,
//           DraftStartDate: x.DraftStartDate,
//           DraftEndDate: x.DraftEndDate,
//           lesson: [
//             { ...x }
//             // [{ ...x }]
//           ]
//         })
//       }
//       else {
//         const lessonIndex = endResp.findIndex((item: any) => item.StepID === x.StepID);
//         let lessonIdAbs = endResp[lessonIndex];// main object
//         const lessonData = endResp[lessonIndex].lesson; //main array


//         lessonData.push({ ...x })
//         endResp[lessonIndex] = {
//           LessonID: lessonIdAbs.LessonID,
//           l_name: lessonIdAbs.l_name,
//           Project: x.Project,
//           StepID: x.StepID,
//           StepName: x.StepName,
//           lesson: lessonData
//         }
//       }
//     })



//     return endResp;

//   }


//   //  :::: first add draft from step manually::::
//   const revManager = (rev_loogs: any[]) => {


//     if (rev_loogs.length == 0) {
//       return [[]]

//     }

//     const validResponses: { [key: string]: string } = {
//       "Review": "Endorsed",
//       "Initial Edit": "Edited",
//       "Assemble": "Assembled",
//       "Publish": "Publish ready",
//       "Sign-off": "Signed Off",
//       "Distribute": "Signed Off",

//     };

//     const resStatusChecker = (rev_item: any) => {
//       // if (rev_item.request == "Draft") {
//       //   return true;
//       // }

//       return validResponses[rev_item.request] === rev_item.response;
//     };

//     var finalArray: any[] = []
//     let currentlistedArray = 0;
//     let count = 1;
//     let req_type = "";
//     let res_type: boolean = false;


//     rev_loogs.forEach(function (rev_item: any) {

//       if (count == 1) {
//         finalArray[currentlistedArray] = [{ ...rev_item, count_arr: count }]
//       } else if (res_type === true) {
//         finalArray[currentlistedArray].push({ ...rev_item, count_arr: count })
//       }
//       else {
//         finalArray[currentlistedArray + 1] = [{ ...rev_item, count_arr: count }]
//         currentlistedArray = currentlistedArray + 1;
//       }
//       count = count + 1;
//       req_type = rev_item.request;
//       res_type = resStatusChecker(rev_item);
//     });

//     return finalArray;
//     console.log(finalArray, "finalArray")
//   }



//   //Function-Section Ends
//   useEffect(() => {
//     if (
//       adpAutoSave &&
//       adpEditFlag &&
//       adpData.some((data) => data.dateError == true) == false
//     ) {
//       setTimeout(() => {
//         newDataFlag
//           ? document.getElementById("adpbtnSave").click()
//           : document.getElementById("adpbtnUpdate").click();
//       }, 300000);
//     }
//   }, [adpAutoSave]);

//   useEffect(() => {
//     setAdpLoader("startUpLoader");
//     setAdpLoader2("startUpLoader");
//     getActivityPlanItem();
//     getActivityPBData();
//     adpGetCurrentUserDetails();
//   }, [adpReRender]);
//   return (
//     <>
//       <div style={{ padding: "5px 15px" }}>
//         {adpLoader2 == "startUpLoader" ? <CustomLoader /> : null}
//         {/* Header-Section Starts */}
//         <div className={styles.adpHeaderSection} style={{ paddingBottom: "0" }}>
//           {/* Popup-Section Starts */}
//           <div></div>
//           {/* Popup-Section Ends */}
//           <div className={styles.adpHeader} style={{ marginBottom: "15px" }}>
//             <div className={styles.dpTitle}>
//               <Icon
//                 iconName="NavigateBack"
//                 className={adpIconStyleClass.navArrow}
//                 onClick={() => {
//                   adpAutoSave
//                     ? confirm(
//                       "You have unsaved changes, are you sure you want to leave?"
//                     )
//                       ? props.handleclick("ActivityPlan", null, "adp")
//                       : null
//                     : props.handleclick("ActivityPlan", null, "adp");
//                 }}
//               />
//               <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
//                 Activity planner
//               </Label>
//             </div>
//             {/* <div style={{ display: "flex" }}>
//               <Persona
//                 size={PersonaSize.size32}
//                 presence={PersonaPresence.none}
//                 imageUrl={
//                   "/_layouts/15/userphoto.aspx?size=S&username=" +
//                   `${
//                     activtyPlanItem.length > 0
//                       ? activtyPlanItem[0]["DeveloperDetails"].email
//                       : ""
//                   }`
//                 }
//               />
//               <Label>
//                 {activtyPlanItem.length > 0
//                   ? activtyPlanItem[0]["DeveloperDetails"].name
//                   : ""}
//               </Label>
//             </div> */}
//           </div>
//           <div
//             style={{
//               display: "flex",
//               justifyContent: "space-between",
//               flexWrap: "wrap",
//             }}
//           >
//             <div
//               className={styles.adpHeaderDetails}
//               style={{ marginLeft: "-10px" }}
//             >
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Project :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0
//                     ? activtyPlanItem[0].Project +
//                     " " +
//                     activtyPlanItem[0].ProjectVersion
//                     : ""}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Product :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0
//                     ? activtyPlanItem[0].Product +
//                     " " +
//                     activtyPlanItem[0].ProductVersion
//                     : ""}
//                 </Label>
//               </div>
//               <div>
//                 {/* <Label style={{ marginRight: 5 }}>
//                   Number of records :{" "}
//                   <span style={{ color: "#038387" }}>
//                     {adpActivityResponseData.length}
//                   </span>
//                 </Label> */}
//               </div>
//               {/* <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Status :</Label>
//                 <Label style={{ color: "#038387", marginRight: "-25px" }}>
//                   {overallStatus()}
//                 </Label>
//               </div> *
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Type :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Project :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>AH/PH :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {overallActualHours()}/{overallPlannedHours()}
//                 </Label>
//               </div> */}
//             </div>

//           </div>
//           <div
//             style={{
//               display: "flex",
//               justifyContent: "space-between",
//               flexWrap: "wrap",
//             }}
//           >
//             <div
//               className={styles.adpHeaderDetails}
//               style={{ marginLeft: "-10px" }}
//             >
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>BA :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {annualPlanData.ba}

//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Term :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {annualPlanData.term}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Hours :</Label>
//                 <Label style={{ color: "#038387" }}> {totalHours}
//                   {/* {activtyPlanItem.length > 0
//                     ? activtyPlanItem[0].Project :''
//                    } */}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Start :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {Tabledata.length > 0
//                     ? Tabledata[0][0].Start : ''
//                   }
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>End :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {Tabledata.length > 0
//                     ? Tabledata[Tabledata.length - 1][Tabledata[Tabledata.length - 1].length - 1].End
//                     : ''}
//                 </Label>
//               </div>

//               {/* <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Status :</Label>
//                 <Label style={{ color: "#038387", marginRight: "-25px" }}>
//                   {overallStatus()}
//                 </Label>
//               </div> *
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Type :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0 ? activtyPlanItem[0].Types : ""}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>Project :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {activtyPlanItem.length > 0 ? activtyPlanItem[0].Project : ""}
//                 </Label>
//               </div>
//               <div style={{ margin: "0 25px 0 10px" }}>
//                 <Label style={{ marginRight: 5 }}>AH/PH :</Label>
//                 <Label style={{ color: "#038387" }}>
//                   {overallActualHours()}/{overallPlannedHours()}
//                 </Label>
//               </div> */}
//             </div>
//             <div style={{ display: "flex" }}>
//               <div
//                 style={{
//                   display: "flex",
//                   justifyContent: "flex-end",
//                   marginTop: 2,
//                   marginRight: 20,
//                 }}
//               >

//               </div>
//               {/* <Label
//                 onClick={() => {
//                   generateExcel();
//                 }}
//                 style={{
//                   backgroundColor: "#EBEBEB",
//                   padding: "0 15px",
//                   cursor: "pointer",
//                   fontSize: "12px",
//                   display: "flex",
//                   alignItems: "center",
//                   justifyContent: "center",
//                   borderRadius: "3px",
//                   color: "#1D6F42",
//                   height: 34,
//                   marginRight: 10,
//                 }}
//               >
//                 <Icon
//                   style={{
//                     color: "#1D6F42",
//                     marginRight: 5,
//                   }}
//                   iconName="ExcelDocument"
//                   className={adpIconStyleClass.export}
//                 />
//                 Export as XLS
//               </Label> */}
//               {adpEditFlag ? (
//                 <PrimaryButton
//                   className={adpbuttonStyleClass.buttonPrimary}
//                   iconProps={cancelIcon}
//                   onClick={() => {
//                     setAdpEDSort("");
//                     setAdpSDSort("");
//                     setAdpEditFlag(false);
//                     setAdpAutoSave(false);
//                     // setAdpData([...adpMasterData]);
//                     // groups([...adpMasterData]);
//                     // setAdpActivityResponseData([...adpMasterData]);
//                     // setFinalStepConst(datauniqueArr)
//                     setFinalStep(finalStepConst)
//                     // setAdpFilters({ ...adpFilterKeys });
//                   }}
//                 >
//                   Cancel
//                 </PrimaryButton>
//               ) : (
//                 <PrimaryButton
//                   className={adpbuttonStyleClass.buttonPrimary}
//                   iconProps={editIcon}
//                   onClick={() => {
//                     setAdpEditFlag(true);
//                     setAdpAutoSave(true);
//                   }}
//                 >
//                   Edit
//                 </PrimaryButton>
//               )}

//               <PrimaryButton
//                 id="adpbtnUpdate"
//                 iconProps={saveIcon}
//                 className={
//                   adpEditFlag &&
//                     adpData.some(
//                       (data) => data.dateError == true || data.PHError == true
//                     ) == false
//                     ? adpbuttonStyleClass.buttonSecondary
//                     : styles.adpSaveBtnDisabled
//                 }
//                 disabled={
//                   adpEditFlag &&
//                     adpData.some(
//                       (data) => data.dateError == true || data.PHError == true
//                     ) == false
//                     ? false
//                     : true
//                 }
//                 onClick={() => {
//                   console.log(diliveryPlanNeedToBeUpdate, "updated h???")
//                   adpUpdateItem()

//                 }}
//               >
//                 {adpLoader == "updateLoader" ? <Spinner /> : <>Save</>}
//               </PrimaryButton>
//               {/* )} */}
//               <Icon
//                 iconName="Link12"
//                 className={adpIconStyleClass.link}
//                 onClick={() => {
//                   adpAutoSave
//                     ? confirm(
//                       "You have unsaved changes, are you sure you want to leave?"
//                     )
//                       ? props.handleclick(
//                         "ActivityProductionBoard",
//                         activityPlan_ID,
//                         "ADP"
//                       )
//                       : null
//                     : props.handleclick(
//                       "ActivityProductionBoard",
//                       activityPlan_ID,
//                       "ADP"
//                     );
//                 }}
//               />
//             </div>
//           </div>
//           {/* Header-Section Ends */}
//           {/* Filter-Section Starts */}
//           <div>
//             <div
//               style={{
//                 display: "flex",
//                 justifyContent: "space-between",
//                 marginTop: "-5px",
//                 marginBottom: "10px",
//                 flexWrap: "wrap",
//               }}
//             >

//             </div>
//           </div>
//           {/* Filter-Section Ends */}
//         </div>

//         {/* Body-Section Starts */}

//         <div>
//           {/* dont remove */}
//           {/* <input
//             id="forFocus"
//             type="text"
//             style={{
//               width: 0,
//               height: 0,
//               border: "none",
//               position: "absolute",
//               top: 0,
//               left: 0,
//               padding: 0,
//             }}
//           /> */}
//         </div>
//         <div
//           className={styles.scrollTop}
//           onClick={() => {
//             document.querySelector("#forFocus")["focus"]();
//           }}
//         >
//           <Icon iconName="Up" style={{ color: "#fff" }} />
//         </div>
//         <div>
//           {/* Table-Section Starts */}




//           {/* <div className="lessonDiv">
//             {Open?.some((arrVal) => index == arrVal) ? (
//               <svg xmlns="http://www.w3.org/2000/svg" style={{
//                 width: '54px',

//                 height: '16px', cursor: 'pointer'
//               }} viewBox="0 0 24 24" onClick={() => show(index)}>
//                 <path d="M7 14l5-5 5 5z" />
//               </svg>
//             ) : (
//               <svg xmlns="http://www.w3.org/2000/svg" style={{
//                 width: '24px',

//                 height: '16px', cursor: 'pointer'
//               }} viewBox="0 0 24 24" onClick={() => show(index)}>
//                 <path d="M12 14l5-5H7z" />
//               </svg>)}


//               Lesson {data[0]?.LessonID + '(' + data.length + ')' + ' '}
//                {((data.filter(task => task.Status === 'Completed').length / data.length) * 100)+"% Complete"}
//             </div> */}
//           <div style={{ overflowX: "auto" }} id="style-1" className="MyTable_2">
//             {/* {Open?.some((arrVal) => index == arrVal) && <>  */}
//             <table className="table table-bordered fixed-width-table">

//               <tr style={{ position: "sticky", top: "0", zIndex: "6" }}>
//                 <th className="Title text-center-do" style={{ width: "14%" }} >Task</th>
//                 <th className="Title text-center-do" style={{ width: "8%" }}>Activity</th>
//                 <th className="Title text-center-do" style={{ width: "12%" }}>Step</th>



//                 {StepsArrayCustomized.map
//                   ((stepInfo, i) =>
//                   (
//                     <th className="Title  text-center-do" style={{ width: "11%" }} colSpan={2}>{stepInfo.name}
//                     </th>))
//                 }


//               </tr>


//               {finalStep.length > 0 ? finalStep.map((data, index) => {

//                 return <>

//                   {data.lesson.map((nestedArrays, ind) => {
//                     return <>

//                       <tr>
//                         <td className="typeData" > {ind == 0 && data.Project}</td>
//                         <td >{ind == 0 && data.l_name}</td>
//                         <td >{data.StepName}   </td>

//                         <>


//                           <td>
//                             {/* {nestedArrays.find((obj) => obj.request === "Draft") ? nestedArrays.find((obj) => obj.request === "Draft").Start : ''} */}


//                             {ind == 0 && adpEditFlag ? (
//                               <>
//                                 <DatePicker
//                                   placeholder="Select a start date"
//                                   formatDate={dateFormater}
//                                   // minDate={new Date(item.Start)}
//                                   // maxDate={new Date(item.End)}
//                                   styles={{
//                                     textField: {
//                                       transform: "translateY(3px)",
//                                       selectors: {
//                                         ".ms-TextField-fieldGroup": {
//                                           borderColor: "#000",
//                                           borderRadius: 4,
//                                           border: "1px solid",
//                                           height: 23,
//                                           input: {
//                                             borderRadius: 4,
//                                           },
//                                         },
//                                         ".ms-TextField-field": {
//                                           color: "#000",
//                                         },
//                                         ".ms-DatePicker-event--without-label": {
//                                           color: "#000",
//                                           paddingTop: 3,
//                                         },
//                                       },
//                                     },
//                                     readOnlyTextField: {
//                                       lineHeight: 22,
//                                     },
//                                   }}
//                                   value={
//                                     data.DraftStartDate
//                                       ? new Date(

//                                         moment(data.DraftStartDate, DateListFormat).format(DatePickerFormat)

//                                       )
//                                       : new Date()
//                                   }
//                                   onSelectDate={(value: any) => {
//                                     console.log("----")

//                                     // const inputDate = new Date(value);
//                                     // const convertedDate = convertDateFormat(inputDate);
//                                     // console.log(convertedDate); 
//                                     const inputDate = new Date(value);
//                                     const customFormat = "DD/MM/YYYY";
//                                     const convertedDate = convertToCustomFormat(inputDate, customFormat);
//                                     console.log(value, "change date", convertedDate)
//                                     let finalStepTemp = finalStep;

//                                     let respData = finalStepTemp[index]
//                                     if (convertedDate) { respData.DraftStartDate = convertedDate }

//                                     const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
//                                     const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.StepID === respData.StepID)
//                                     if (dilIndex == -1) {
//                                       diliveryPlanNeedToBeUpdateTemp.push(respData)
//                                     } else {
//                                       diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
//                                     }
//                                     setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update

//                                     console.log(finalStepTemp, "after date")
//                                     finalStepTemp[index] = respData
//                                     setFinalStep(finalStepTemp)
//                                     setreRenderState(!reRenderState)
//                                   }}
//                                 />
//                               </>
//                             ) : (
//                               <>
//                                 {ind == 0 && (<>{data.DraftStartDate}</>)}
//                               </>
//                             )}
//                           </td>
//                           <td  >
//                             {/* {nestedArrays.find((obj) => obj.request === "Draft") ? nestedArrays.find((obj) => obj.request === "Draft").End : ''} */}
//                             {ind == 0 && adpEditFlag ? (
//                               <>
//                                 <DatePicker
//                                   placeholder="Select a start date"
//                                   formatDate={dateFormater}
//                                   // minDate={new Date(item.Start)}
//                                   // maxDate={new Date(item.End)}
//                                   styles={{
//                                     textField: {
//                                       transform: "translateY(3px)",
//                                       selectors: {
//                                         ".ms-TextField-fieldGroup": {
//                                           borderColor: "#000",
//                                           borderRadius: 4,
//                                           border: "1px solid",
//                                           height: 23,
//                                           input: {
//                                             borderRadius: 4,
//                                           },
//                                         },
//                                         ".ms-TextField-field": {
//                                           color: "#000",
//                                         },
//                                         ".ms-DatePicker-event--without-label": {
//                                           color: "#000",
//                                           paddingTop: 3,
//                                         },
//                                       },
//                                     },
//                                     readOnlyTextField: {
//                                       lineHeight: 22,
//                                     },
//                                   }}
//                                   value={
//                                     data.DraftEndDate
//                                       ? new Date(

//                                         moment(data.DraftEndDate, DateListFormat).format(DatePickerFormat)

//                                       )
//                                       : new Date()
//                                   }
//                                   onSelectDate={(value: any) => {
//                                     console.log("----")

//                                     // const inputDate = new Date(value);
//                                     // const convertedDate = convertDateFormat(inputDate);
//                                     // console.log(convertedDate); 
//                                     const inputDate = new Date(value);
//                                     const customFormat = "DD/MM/YYYY";
//                                     const convertedDate = convertToCustomFormat(inputDate, customFormat);
//                                     console.log(value, "change date", convertedDate)
//                                     let finalStepTemp = finalStep;

//                                     let respData = finalStepTemp[index]
//                                     if (convertedDate) { respData.DraftEndDate = convertedDate }

//                                     const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
//                                     const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.StepID === respData.StepID)
//                                     if (dilIndex == -1) {
//                                       diliveryPlanNeedToBeUpdateTemp.push(respData)
//                                     } else {
//                                       diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
//                                     }
//                                     setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update

//                                     console.log(finalStepTemp, "after date")
//                                     finalStepTemp[index] = respData
//                                     setFinalStep(finalStepTemp)
//                                     setreRenderState(!reRenderState)
//                                   }}
//                                 />
//                               </>
//                             ) : (
//                               <>
//                                 {<>{data.DraftEndDate}</>}
//                               </>
//                             )}
//                           </td>
//                           <td

//                           >
//                             {nestedArrays.find((obj) => obj.request === "Review") ? dateFormater(nestedArrays.find((obj) => obj.request === "Review").Start) : ''}

//                           </td>
//                           <td
//                           >
//                             {nestedArrays.find((obj) => obj.request === "Review") ? dateFormater(nestedArrays.find((obj) => obj.request === "Review").End) : ''}

//                           </td>

//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Initial Edit") ? dateFormater(nestedArrays.find((obj) => obj.request === "Initial Edit").Start) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Initial Edit") ? dateFormater(nestedArrays.find((obj) => obj.request === "Initial Edit").End) : ''}

//                           </td>

//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Assemble") ? dateFormater(nestedArrays.find((obj) => obj.request === "Assemble").Start) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Assemble") ? dateFormater(nestedArrays.find((obj) => obj.request === "Assemble").End) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Sign-off") ? nestedArrays.find((obj) => obj.request === "Sign-off").Start : nestedArrays.find((obj) => obj.request === "Publish") ? dateFormater(nestedArrays.find((obj) => obj.request === "Publish").Start) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Sign-off") ? nestedArrays.find((obj) => obj.request === "Sign-off").Start : nestedArrays.find((obj) => obj.request === "Publish") ? dateFormater(nestedArrays.find((obj) => obj.request === "Publish").Start) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Distribute") ? dateFormater(nestedArrays.find((obj) => obj.request === "Distribute").Start) : ''}

//                           </td>
//                           <td  >
//                             {nestedArrays.find((obj) => obj.request === "Distribute") ? dateFormater(nestedArrays.find((obj) => obj.request === "Distribute").End) : ''}

//                           </td>


//                         </>


//                       </tr>

//                       <tr>
//                         <td className="typeData"  >
//                         </td>
//                         <td  > </td>
//                         <td  > </td>

//                         <>
//                           <td colSpan={2}  >




//                             {ind == 0 &&
//                               adpEditFlag ? (
//                               <>
//                                 <NormalPeoplePicker
//                                   styles={{
//                                     root: {
//                                       selectors: {
//                                         ".ms-SelectionZone": {
//                                           height: 24,
//                                         },
//                                         ".ms-BasePicker-text": {
//                                           height: 24,
//                                           padding: 1,
//                                           border: "1px solid #000",
//                                           borderRadius: 4,
//                                           marginTop: -6,
//                                           marginRight: 20,
//                                         },
//                                       },
//                                     },
//                                   }}
//                                   onResolveSuggestions={GetUserDetails}
//                                   itemLimit={1}
//                                   selectedItems={allPeoples.filter((people) => {
//                                     return (
//                                       people.ID == data?.Developer?.id ? data?.Developer?.id : null
//                                     );
//                                   })}
//                                   onChange={(selectedUser) => {
//                                     console.log("change image", selectedUser)
//                                     let finalStepTemp = finalStep;

//                                     let respData = finalStepTemp[index]
//                                     console.log(respData, "before")
//                                     if (selectedUser[0]) {


//                                       respData.Developer = {
//                                         id: selectedUser[0]["ID"],
//                                         email: selectedUser[0]["secondaryText"],
//                                         name: selectedUser[0]["text"]
//                                       }
//                                       // respData.Dev = selectedUser[0]["text"]
//                                       // respData.FromEmail = selectedUser[0]["secondaryText"]
//                                       // respData.ToEmail = selectedUser[0]["secondaryText"]
//                                     } else {



//                                       respData.Developer = {
//                                         id: undefined,
//                                         email: undefined,
//                                         name: undefined
//                                       }
//                                       // respData.Dev = undefined
//                                       // respData.FromEmail = undefined
//                                       // respData.ToEmail = undefined

//                                     }

//                                     const diliveryPlanNeedToBeUpdateTemp = diliveryPlanNeedToBeUpdate;
//                                     const dilIndex = diliveryPlanNeedToBeUpdateTemp.findIndex((obj) => obj.StepID === respData.StepID)
//                                     if (dilIndex == -1) {
//                                       diliveryPlanNeedToBeUpdateTemp.push(respData)
//                                     } else {
//                                       diliveryPlanNeedToBeUpdateTemp[dilIndex] = respData
//                                     }
//                                     setDiliveryPlanNeedToBeUpdate(diliveryPlanNeedToBeUpdateTemp) //string data of selected draft including activity dilivery plan id which will be use latter to update
//                                     finalStepTemp[index] = respData
//                                     console.log(ind, "nested index")
//                                     console.log(respData, "after")
//                                     // console.log(finalStepTemp, "1234567890")
//                                     setFinalStep(finalStepTemp)
//                                     setreRenderState(!reRenderState)



//                                     // adpActivityResponseHandler(
//                                     //   1,
//                                     //   "Developer",
//                                     //   selectedUser[0] ? selectedUser[0]["ID"] : null
//                                     // ); 
//                                   }}
//                                 />
//                               </>
//                             ) : (
//                               <>
//                                 {ind == 0 && newDataFlag ? (
//                                   <>
//                                     <TooltipHost
//                                       id="myPersonaTooltip"
//                                       style={{ display: "flex", justifyContent: "center" }}
//                                     >
//                                       <Persona
//                                         size={PersonaSize.size32}
//                                         presence={PersonaPresence.none}
//                                         imageUrl={

//                                           `/_layouts/15/userphoto.aspx?size=S&username=${data.Developer.email}`
//                                         }
//                                       />
//                                     </TooltipHost>

//                                   </>
//                                 ) : ind == 0 && data?.Developer?.email ? (
//                                   <>
//                                     <TooltipHost
//                                       id="myPersonaTooltip"
//                                       style={{ display: "flex", justifyContent: "center" }}
//                                     >
//                                       <Persona
//                                         size={PersonaSize.size32}
//                                         presence={PersonaPresence.none}
//                                         imageUrl={
//                                           `/_layouts/15/userphoto.aspx?size=S&username=${data.Developer.email}`
//                                         }
//                                       />
//                                     </TooltipHost>

//                                   </>
//                                 ) : (
//                                   ""
//                                 )}
//                               </>
//                             )
//                             }



//                             {/* 
//                             {nestedArrays.find((obj) => obj?.request === "Draft") ? 
//                               <>
//                                 <TooltipHost
//                                   id="myPersonaTooltip"
//                                   style={{ display: "flex", justifyContent: "center" }}
//                                 >
//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Draft")?.FromEmail}`
//                                     }
//                                   />
//                                 </TooltipHost>

//                               </>
 
//                               : ''}

// */}


//                           </td>

//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Review") ?

//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Review")?.Dev}
//                                   id="myPersonaTooltip"
//                                 >
//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Review")?.FromEmail}`
//                                     }
//                                   />
//                                 </TooltipHost>

//                               </>


//                               : ''}

//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Review") ?
//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Review")?.client}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Review")?.ToEmail}`
//                                     }
//                                   />

//                                 </TooltipHost>
//                               </>


//                               : ''}

//                           </td>


//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Initial Edit") ?
//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Initial Edit")?.Dev}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Initial Edit")?.FromEmail}`
//                                     }
//                                   />
//                                 </TooltipHost>
//                               </>



//                               : ''}

//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Initial Edit") ?

//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Initial Edit")?.client}
//                                   id="myPersonaTooltip"
//                                 >
//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Initial Edit")?.ToEmail}`
//                                     }
//                                   /></TooltipHost></>




//                               : ''}

//                           </td>



//                           <td  >


//                             {nestedArrays.find((obj) => obj?.request === "Assemble") ?
//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Assemble")?.Dev}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Assemble")?.FromEmail}`
//                                     }
//                                   />
//                                 </TooltipHost></>



//                               : ''}

//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Assemble") ?

//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Assemble")?.client}
//                                   id="myPersonaTooltip"
//                                 >
//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Assemble")?.ToEmail}`
//                                     }
//                                   />
//                                 </TooltipHost></>



//                               : ''}

//                           </td>       <td  >

//                             {nestedArrays.find((obj) => obj?.request === "Sign-off") ?

//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Sign-off")?.Dev}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Sign-off")?.FromEmail}`
//                                     }
//                                   /> </TooltipHost> </> : nestedArrays.find((obj) => obj.request === "Publish") ?
//                                 <>
//                                   <TooltipHost
//                                     content={nestedArrays.find((obj) => obj?.request === "Publish")?.Dev}
//                                     id="myPersonaTooltip"
//                                   >

//                                     <Persona
//                                       size={PersonaSize.size32}
//                                       presence={PersonaPresence.none}
//                                       imageUrl={
//                                         "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                         `${nestedArrays.find((obj) => obj?.request === "Publish")?.FromEmail}`
//                                       }
//                                     /></TooltipHost> </> : ''}

//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Sign-off") ?
//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Sign-off")?.client}
//                                   id="myPersonaTooltip"
//                                 >
//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Sign-off")?.ToEmail}`
//                                     }
//                                   /> </TooltipHost>
//                               </> : nestedArrays.find((obj) => obj.request === "Publish") ?

//                                 <>
//                                   <TooltipHost
//                                     content={nestedArrays.find((obj) => obj?.request === "Publish")?.client}
//                                     id="myPersonaTooltip"
//                                   >
//                                     <Persona
//                                       size={PersonaSize.size32}
//                                       presence={PersonaPresence.none}
//                                       imageUrl={
//                                         "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                         `${nestedArrays.find((obj) => obj?.request === "Publish")?.ToEmail}`
//                                       }
//                                     /> </TooltipHost>
//                                 </> : ''}

//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Distribute") ?
//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Distribute")?.Dev}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Distribute")?.FromEmail}`
//                                     }
//                                   />

//                                 </TooltipHost>
//                               </>
//                               : ''}
//                           </td>
//                           <td  >



//                             {nestedArrays.find((obj) => obj?.request === "Distribute") ?

//                               <>
//                                 <TooltipHost
//                                   content={nestedArrays.find((obj) => obj?.request === "Distribute")?.client}
//                                   id="myPersonaTooltip"
//                                 >

//                                   <Persona
//                                     size={PersonaSize.size32}
//                                     presence={PersonaPresence.none}
//                                     imageUrl={
//                                       "/_layouts/15/userphoto.aspx?size=S&username=" +
//                                       `${nestedArrays.find((obj) => obj?.request === "Distribute")?.ToEmail}`
//                                     }
//                                   />

//                                 </TooltipHost>
//                               </>

//                               : ''}
//                           </td>


//                         </>


//                       </tr>

//                     </>
//                   })}





//                 </>

//               }) : <tr>
//                 <td className="typeData" colSpan={14} width="100%">      <Label style={{ color: "#2392B2", textAlign: "center" }}>No Data Found !!!</Label> </td>



//               </tr>
//               }


//             </table>
//             {/* </>} */}
//           </div>





//           {/* Table-Section Ends */}
//         </div>

//         <div>
//           <Modal isOpen={AdpConfirmationPopup.condition} isBlocking={true}>
//             <div
//               style={{
//                 display: "flex",
//                 justifyContent: "center",
//                 alignItems: "center",
//                 marginTop: "30px",
//                 width: "450px",
//               }}
//             >
//               <div
//                 style={{
//                   display: "flex",
//                   alignItems: "center",
//                   justifyContent: "flex-Start",
//                   flexDirection: "column",
//                   marginBottom: "10px",
//                 }}
//               >
//                 <Label className={styles.deletePopupTitle}>Confirmation</Label>
//                 <Label
//                   style={{
//                     padding: "5px 20px",
//                   }}
//                   className={styles.deletePopupDesc}
//                 >
//                   Are you sure want to mark as completed?
//                 </Label>
//               </div>
//             </div>
//             <div className={styles.apDeletePopupBtnSection}>
//               <button
//                 onClick={(_) => {
//                   setAdpConfirmationPopup({ condition: false, isNew: false });
//                   // saveDPData();
//                   setAdpLoader("startUpLoader");
//                   AdpConfirmationPopup.isNew ? adpAddItem() : null
//                 }}
//                 className={styles.apDeletePopupYesBtn}
//               >
//                 Yes
//               </button>
//               <button
//                 onClick={(_) => {
//                   setAdpConfirmationPopup({ condition: false, isNew: false });
//                 }}
//                 className={styles.apDeletePopupNoBtn}
//               >
//                 No
//               </button>
//             </div>
//           </Modal>
//         </div>

//         {/* Body-Section Ends */}
//       </div>
//     </>
//   );
// };

// export default ActivityDeliveryPlan;
