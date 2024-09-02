import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  PrimaryButton,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";

import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { arraysEqual } from "office-ui-fabric-react/lib/Utilities";

const PL_Resources = (props: any) => {
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;

  const currentBA = props.PLBAObject.BA;
  const currentSubject = props.PLBAObject.Subject;
  const currentProduct = props.PLBAObject.Product;
  const currentProductVersion = props.PLBAObject.ProductVersion;
  const currentProject = props.PLBAObject.Project;
  const currentProjectVersion = props.PLBAObject.ProjectVersion;

  let currentpage = 1;
  let totalPageItems = 10;
  let sortPLDataArr = [];
  let sortPLFilterArr = [];

  const PLFilterKey = {
    Lesson: "",
    Title: "",
    Status: "All",
  };
  const PLDrpDownOptns = {
    TypeOfProject: [{ key: "All", text: "All" }],
    Status: [{ key: "All", text: "All" }],
  };

  const _PLAPRColumns = [
    {
      key: "Column1",
      name: "Resource",
      fieldName: "Lesson",
      minWidth: 300,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column2",
      name: "Activity Plan name / Template",
      fieldName: "Title",
      minWidth: 300,
      maxWidth: 400,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column3",
      name: "Status",
      fieldName: "Status",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={PLStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Under development" ? (
            <div className={PLStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "Not started" ? (
            <div className={PLStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : (
            <div className={PLStatusStyleClass.Onhold}>{item.Status}</div>
          )}
        </>
      ),
    },
    {
      key: "Column4",
      name: "Completion",
      fieldName: "Completion",
      minWidth: 150,
      maxWidth: 210,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <div style={{ width: 200, position: "relative" }}>
          {item.Completion >= 75 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.completed}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 50 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.scheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 25 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.onSchedule}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : item.Completion >= 0 ? (
            <div className={PLCompletionStyleClass.default}>
              <span className={PLCompletionStyleClass.percentageText}>
                {Math.round(item.Completion)}%
              </span>{" "}
              <div
                style={{ width: `${item.Completion}%` }}
                className={PLCompletionStyleClass.behindScheduled}
              >
                {/* {item.Completion} */}
              </div>
            </div>
          ) : (
            ""
          )}
        </div>
      ),
    },
  ];
  const PLSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const PLActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 200,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "2px solid #038387",
      borderRadius: "4px",
      marginTop: "3px",
    },
    field: { fontWeight: 600, color: "#038387" },
    icon: { fontSize: 14, color: "#038387" },
  };
  const PLlabelStyles = mergeStyleSets({
    titleLabel: [
      {
        color: "#676767",
        fontSize: "14px",
        marginRight: "10px",
        fontWeight: "400",
      },
    ],
    labelValue: [
      {
        color: "#0882A5",
        fontSize: "14px",
        marginRight: "10px",
      },
    ],
    inputLabels: [
      {
        color: "#323130",
        fontSize: "13px",
      },
    ],
    navLabels: [
      {
        color: "#2392B2",
        fontSize: "16px",
        cursor: "pointer",
      },
    ],
    navViewLabels: [
      {
        fontSize: "16px",
        cursor: "pointer",
      },
    ],
    ErrorLabel: [
      {
        marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#D0342C",
        fontSize: "13px",
      },
    ],
    NORLabel: [
      {
        marginTop: 28,
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });
  const PLProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 15px 0 0",
  });
  const PLIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const PLIconStyleClass = mergeStyleSets({
    link: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
        marginLeft: "4px",
      },
    ],
    rightArrow: [
      { color: "#2392B2", marginRight: 10, fontSize: 20 },
      PLIconStyle,
    ],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 31,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: 3,
        marginRight: 12,
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
        marginRight: 5,
      },
    ],
  });
  const PLStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: 120,
  });
  const PLStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PLStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PLStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "rgb(0 191 198)",
        backgroundColor: "rgb(210 241 241)",
      },
      PLStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PLStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      PLStatusStyle,
    ],
  });
  const PLCompletionStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    height: 17,
    width: 200,
  });
  const PLCompletionStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      PLCompletionStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      PLCompletionStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        color: "#9C9C00; ",
        backgroundColor: "#EEEEAE",
      },
      PLCompletionStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      PLCompletionStyle,
    ],
    default: [
      {
        fontWeight: "600",
        position: "relative",
        backgroundColor: "#edebe9",
      },
      PLCompletionStyle,
    ],
    percentageText: [
      {
        position: "absolute !important",
        left: "50%",
        top: "50%",
        transform: "translate(-50%,-50%)",
        color: "#555",
      },
    ],
  });
  const PLDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: 4,
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const PLActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 200,
      marginTop: 5,
      marginRight: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
      fontWeight: 600,
    },
    callout: {
      maxHeight: "300px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };

  // Use State
  const [PLAPRReRender, setPLAPRReRender] = useState(false);
  const [PLAPRMaster, setPLAPRMaster] = useState([]);
  const [PLAPRData, setPLAPRData] = useState([]);
  const [PLAPRFilData, setPLAPRFilData] = useState([]);
  const [PLAPRDisplay, setPLAPRDisplay] = useState([]);
  const [PLAPRFilterKey, setPLAPRFilterKey] = useState(PLFilterKey);
  const [PLAPRDrpDownOptns, setPLAPRDrpDownOptns] = useState(PLDrpDownOptns);
  const [PLAPRLoader, setPLAPRLoader] = useState("noLoader");
  const [PLAPRColumns, setPLAPRColumns] = useState(_PLAPRColumns);
  const [PLAPRcurPage, setPLAPRCurPage] = useState(currentpage);

  const getActivityPlanMaster = () => {
    let _activityPlanArr = [];
    sharepointWeb.lists
      .getByTitle("Activity Plan")
      .items.filter(
        `Product eq '${currentProduct}' and Project eq '${currentProject}'`
      )
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        let tempLessonArr = [];
        items.forEach((item) => {
          let PrjVersion = item.ProjectVersion ? item.ProjectVersion : "V1";
          let PrdVersion = item.ProductVersion ? item.ProductVersion : "V1";

          if (
            PrjVersion == currentProjectVersion &&
            PrdVersion == currentProductVersion
          ) {
            // _activityPlanArr.push({
            //   ID: item.ID,
            //   Types: item.Types ? item.Types : "",
            //   Area: item.Area ? item.Area : "",
            //   ActivityPlanName: item.ActivityPlanName
            //     ? item.ActivityPlanName
            //     : "",
            //   Template: item.Title ? item.Title : "",
            //   Status: item.Status ? item.Status : "",
            // });

            let lessons_str_to_arr = item.Lessons
              ? item.Lessons.split(";")
              : null;

            if (lessons_str_to_arr) {
              lessons_str_to_arr.forEach((lesson) => {
                tempLessonArr.push({
                  Lesson: lesson.split("~")[1],
                  LessonId: lesson.split("~")[0],
                  ID: item.ID,
                  Types: item.Types ? item.Types : "",
                  Area: item.Area ? item.Area : "",
                  ActivityPlanName: item.ActivityPlanName
                    ? item.ActivityPlanName
                    : "",
                  Template: item.Title ? item.Title : "",
                  Status: item.Status ? item.Status : "",
                });
              });
            }
          }
        });
        // getActivityPlan(_activityPlanArr);
        getActivityPlanner(tempLessonArr);
      })
      .catch((err) => {
        ErrorFunction(err, "getApData");
      });
  };

  const getActivityPlan = (data) => {
    let tempLessonArr = [];
    let count = 0;
    data.forEach((arr) => {
      count++;
      sharepointWeb.lists
        .getByTitle("Activity Plan")
        .items.filter(`ID eq '${arr.ID}'`)

        .top(5000)
        .orderBy("Modified", false)
        .get()
        .then((items) => {
          items.forEach((item) => {
            let lessons_str_to_arr = item.Lessons
              ? item.Lessons.split(";")
              : null;

            if (lessons_str_to_arr) {
              lessons_str_to_arr.forEach((lesson) => {
                if (
                  tempLessonArr.findIndex((arr) => {
                    return arr.Lesson == lesson.split("~")[1];
                  }) == -1 &&
                  lesson.split("~")[1]
                ) {
                  tempLessonArr.push({
                    Lesson: lesson.split("~")[1],
                  });
                }
              });
            }
          });
          if (data.length == count) {
            getActivityPlanner(tempLessonArr);
          }
        })
        .catch((err) => {
          ErrorFunction(err, "getApData");
        });
    });
  };



  const getCamelquery = async (_id) => {


    let response = false;

    let camelQueryXML: string =
      '<View>' +
      "<ViewFields>" +
      "<FieldRef Name='ID'/>" +
      "<FieldRef Name='auditResponseType'/>" +
      "<FieldRef Name='auditRequestType'/>" +
      "</ViewFields>" +
      `<Query>
<Where>
   <And>
      <Eq>
         <FieldRef Name='DeliveryPlanID' />
         <Value Type='Number'>${_id}</Value>
      </Eq>
      <And>
         <Eq>
            <FieldRef Name='auditRequestType' />
            <Value Type='Choice'>Distribute</Value>
         </Eq>
         <Or>
            <Eq>
               <FieldRef Name='auditResponseType' />
               <Value Type='Choice'>Approved</Value>
            </Eq>
            <Eq>
               <FieldRef Name='auditResponseType' />
               <Value Type='Choice'>Completed</Value>
            </Eq>
         </Or>
      </And>
   </And>
</Where>
</Query>` +
      '</View>';


    //sp.web.lists.getByTitle("ProductionBoard").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((productionBoardResponse: any)

    await sharepointWeb.lists
      .getByTitle("Review log").getItemsByCAMLQuery({ 'ViewXml': camelQueryXML }).then((data: any) => {
        if (data.length > 0) {
          response = true;
          console.log(data, "caml");
          // alert("Mil gya")
        }
      });


    if (response) {
      // alert("true")
    }
    return response;
  };




  const getActivityPlanner = (records) => {
    let _activityPlannerArr = [];
    let count = 0;
    const objectttt = []
    if (records.length > 0) {
      records.forEach((record) => {
        sharepointWeb.lists
          .getByTitle("Activity Delivery Plan")
          .items.filter(
            `ActivityPlanID eq '${record.ID}' and LessonID eq '${record.LessonId}'`
          )
          .orderBy("Modified", false)
          .top(5000)
          .get()
          .then((items) => {
            count++;
            //
            // make lop on item against lesson id and check ids of these items have any 
            let isCompleted = false;
            items.map((x) => {
              // 
              const resp = getCamelquery(x.ID)
              if (resp) {
                // alert("true")
                isCompleted = true;
                objectttt.push({
                  record: record.ID,
                  lesson: record.LessonId
                })
              }



            })

            // let completedLength = items.filter((arr) => {
            //   return arr.Status == "Completed";
            // }).length;

            if (isCompleted) {
              console.log({
                Lesson: record.Lesson,
                Title: record.ActivityPlanName
              }, "What the hell is going on")
            }

            _activityPlannerArr.push({
              Lesson: record.Lesson,
              Title: record.ActivityPlanName
                ? record.ActivityPlanName
                : record.Template,
              Status: isCompleted ? "Completed" : "Not started",

              // record.Status == "Completed" ||
              // (items.length == completedLength && items.length > 0)
              //   ? "Completed"
              //   : completedLength != 0 && items.length != 0
              //   ? "Under development"
              //   : "Not started",
              Completion: isCompleted ? 100 : 0,
              // record.Status == "Completed" ||
              // (items.length == completedLength && items.length > 0)
              //   ? 100
              //   : completedLength != 0 && items.length != 0
              //   ? (completedLength / items.length) * 100
              //   : 0,
            });

            // if (
            //   PLDrpDownOptns.Status.findIndex((_projectDrpDwn) => {
            //     return _projectDrpDwn.key == record.Status;
            //   }) == -1 &&
            //   record.Status
            // ) {
            //   PLDrpDownOptns.Status.push({
            //     key: record.Status,
            //     text: record.Status,
            //   });
            // }

            if (records.length == count) {
              PLDrpDownOptns.Status.push(
                {
                  key: "Not started",
                  text: "Not started",
                },
                {
                  key: "Under development",
                  text: "Under development",
                },
                {
                  key: "Completed",
                  text: "Completed",
                }
              );

              setPLAPRDrpDownOptns({ ...PLDrpDownOptns });
              setPLAPRMaster([..._activityPlannerArr]);
              setPLAPRData([..._activityPlannerArr]);
              sortPLDataArr = [..._activityPlannerArr];
              setPLAPRFilData([..._activityPlannerArr]);
              sortPLFilterArr = [..._activityPlannerArr];
              paginate(1, [..._activityPlannerArr]);
              setPLAPRLoader("noLoader");
            }
          });
      });
    } else {
      setPLAPRDrpDownOptns({ ...PLDrpDownOptns });
      setPLAPRMaster([..._activityPlannerArr]);
      setPLAPRData([..._activityPlannerArr]);
      sortPLDataArr = [..._activityPlannerArr];
      setPLAPRFilData([..._activityPlannerArr]);
      sortPLFilterArr = [..._activityPlannerArr];
      paginate(1, [..._activityPlannerArr]);
      setPLAPRLoader("noLoader");
    }
    console.log(objectttt , "-----")
  };

  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setPLAPRDisplay(paginatedItems);
      setPLAPRCurPage(pagenumber);
    } else {
      setPLAPRDisplay([]);
      setPLAPRCurPage(1);
    }
  };
  const PLFilterFunction = (key, value) => {
    let tempFilterKey = { ...PLAPRFilterKey };
    tempFilterKey[key] = value;

    let tempArr = [...PLAPRData];

    if (tempFilterKey.Lesson != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.Lesson.toLowerCase().includes(
          tempFilterKey.Lesson.toLowerCase()
        );
      });
    }
    if (tempFilterKey.Title != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.Title.toLowerCase().includes(
          tempFilterKey.Title.toLowerCase()
        );
      });
    }
    if (tempFilterKey.Status != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Status == tempFilterKey.Status;
      });
    }
    setPLAPRFilData([...tempArr]);
    sortPLFilterArr = [...tempArr];
    paginate(1, [...tempArr]);
    setPLAPRFilterKey({ ...tempFilterKey });
  };
  const ErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    setPLAPRLoader("noLoader");

    // let response = {
    //   ComponentName: "PL_ActivityPlan",
    //   FunctionName: functionName,
    //   ErrorMessage: JSON.stringify(error["message"]),
    //   Title: loggeduseremail,
    // };

    // Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
    //   () => {
    ErrorPopup();
    //   }
    // );
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  // Sorting Function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = _PLAPRColumns;
    const newColumns: IColumn[] = tempapColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newPLDataArr = _copyAndSort(
      sortPLDataArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newPLFilterArr = _copyAndSort(
      sortPLFilterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setPLAPRData([...newPLDataArr]);
    setPLAPRFilData([...newPLFilterArr]);
    paginate(1, newPLFilterArr);
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

  //Use Effect
  useEffect(() => {
    setPLAPRLoader("startUpLoader");
    getActivityPlanMaster();
  }, [PLAPRReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {PLAPRLoader == "startUpLoader" ? (
          <CustomLoader />
        ) : (
          <>
            <div
              style={{
                display: "flex",
                marginBottom: 10,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("BusinessArea");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>Home</Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("Subject");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>{currentBA}</Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("Product");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>
                  {currentSubject}
                </Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
                onClick={() => {
                  props.selectPLFunction("Project");
                }}
              >
                <Label className={PLlabelStyles.navLabels}>
                  {currentProduct + " " + currentProductVersion}
                </Label>
                <Icon
                  iconName="ChevronRight"
                  title="Click to navigate"
                  className={PLIconStyleClass.rightArrow}
                />
              </div>
              <div>
                <Label className={PLlabelStyles.navViewLabels}>
                  {currentProject + " " + currentProjectVersion}
                </Label>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
                marginBottom: 10,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                }}
              >
                <Icon
                  aria-label="ChevronLeftMed"
                  iconName="NavigateBack"
                  className={PLIconStyleClass.ChevronLeftMed}
                  onClick={() => {
                    props.selectPLFunction("Project");
                  }}
                />

                <div className={styles.dpTitle}>
                  <Label style={{ fontSize: 24, padding: 0 }}>Resources</Label>
                </div>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                marginTop: 15,
                justifyContent: "space-between",
              }}
            >
              <div className={styles.Section1}>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>
                    Business area :
                  </Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentBA}
                  </Label>
                </div>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>Subject :</Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentSubject}
                  </Label>
                </div>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>Product :</Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentProduct + " " + currentProductVersion}
                  </Label>
                </div>
                <div className={PLProjectInfo}>
                  <Label className={PLlabelStyles.titleLabel}>
                    Deliverable :
                  </Label>
                  <Label className={PLlabelStyles.labelValue}>
                    {currentProject + " " + currentProjectVersion}
                  </Label>
                </div>
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  flexWrap: "wrap",
                }}
              >
                <div>
                  <Label className={PLlabelStyles.inputLabels}>Resource</Label>
                  <SearchBox
                    styles={
                      PLAPRFilterKey.Lesson
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLAPRFilterKey.Lesson}
                    onChange={(e, value) => {
                      PLFilterFunction("Lesson", value);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>
                    Activity Plan name / Template
                  </Label>
                  <SearchBox
                    styles={
                      PLAPRFilterKey.Title
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLAPRFilterKey.Title}
                    onChange={(e, value) => {
                      PLFilterFunction("Title", value);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>Status</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      PLAPRFilterKey.Status != "All"
                        ? PLActiveDropdownStyles
                        : PLDropdownStyles
                    }
                    options={PLAPRDrpDownOptns.Status}
                    dropdownWidth={"auto"}
                    selectedKey={PLAPRFilterKey.Status}
                    onChange={(e, option: any) => {
                      PLFilterFunction("Status", option["key"]);
                    }}
                  />
                </div>

                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={PLIconStyleClass.refresh}
                    onClick={() => {
                      setPLAPRData([...PLAPRMaster]);
                      sortPLDataArr = [...PLAPRMaster];
                      setPLAPRFilData([...PLAPRMaster]);
                      sortPLFilterArr = [...PLAPRMaster];
                      paginate(1, [...PLAPRMaster]);
                      setPLAPRFilterKey(PLFilterKey);
                      setPLAPRColumns(_PLAPRColumns);
                    }}
                  />
                </div>
              </div>
              <div>
                <Label className={PLlabelStyles.NORLabel}>
                  Number of records:{" "}
                  <b style={{ color: "#038387" }}>{PLAPRFilData.length}</b>
                </Label>
              </div>
            </div>
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={PLAPRDisplay}
                columns={PLAPRColumns}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
                // styles={gridStyles}
                styles={{ root: { width: "100%" } }}
              />
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  margin: "20px 0",
                }}
              >
                {PLAPRFilData.length > 0 ? (
                  <Pagination
                    currentPage={PLAPRcurPage}
                    totalPages={
                      PLAPRFilData.length > 0
                        ? Math.ceil(PLAPRFilData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginate(page, PLAPRFilData);
                    }}
                  />
                ) : (
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "center",
                      marginTop: "15px",
                    }}
                  >
                    <Label style={{ color: "#2392B2" }}>
                      No data Found !!!
                    </Label>
                  </div>
                )}
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
};
export default PL_Resources;
