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

const PL_ActivityPlan = (props: any) => {
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;

  const currentBA = props.PLBAObject.BA;
  const currentProduct = props.PLBAObject.Product;
  const currentProductVersion = props.PLBAObject.ProductVersion;
  const currentProject = props.PLBAObject.Project;
  const currentProjectVersion = props.PLBAObject.ProjectVersion;

  let currentpage = 1;
  let totalPageItems = 10;
  let sortPLDataArr = [];
  let sortPLFilterArr = [];

  const PLFilterKey = {
    Template: "",
    ActivityPlanName: "",
    Types: "All",
    Area: "All",
  };

  const PLDrpDownOptns = {
    Types: [{ key: "All", text: "All" }],
    Area: [{ key: "All", text: "All" }],
  };

  const _PLAPColumns = [
    {
      key: "Column1",
      name: "Types",
      fieldName: "Types",
      minWidth: 150,
      maxWidth: 300,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column2",
      name: "Area",
      fieldName: "Area",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column3",
      name: "Activity plan name",
      fieldName: "ActivityPlanName",
      minWidth: 150,
      maxWidth: 300,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column4",
      name: "Template",
      fieldName: "Template",
      minWidth: 150,
      maxWidth: 300,
      onColumnClick: (ev, column) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column5",
      name: "Link",
      fieldName: "link",
      minWidth: 30,
      maxWidth: 30,

      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={PLIconStyleClass.link}
            onClick={() => {
              props.selectPLFunction(
                "ActivityPlanner",
                "ActivityPlan",
                item.ActivityPlanName ? item.ActivityPlanName : item.Template,
                "",
                item.ID
              );
            }}
          />
        </>
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
  const [PLAPReRender, setPLAPReRender] = useState(false);
  const [PLAPMaster, setPLAPMaster] = useState([]);
  const [PLAPData, setPLAPData] = useState([]);
  const [PLAPFilData, setPLAPFilData] = useState([]);
  const [PLAPDisplay, setPLAPDisplay] = useState([]);
  const [PLAPFilterKey, setPLAPFilterKey] = useState(PLFilterKey);
  const [PLAPDrpDownOptns, setPLAPDrpDownOptns] = useState(PLDrpDownOptns);
  const [PLAPLoader, setPLAPLoader] = useState("noLoader");
  const [PLAPColumns, setPLAPColumns] = useState(_PLAPColumns);
  const [PLAPcurPage, setPLAPCurPage] = useState(currentpage);

  const getActivityPlan = () => {
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
        items.forEach((item) => {
          let PrjVersion = item.ProjectVersion ? item.ProjectVersion : "V1";
          let PrdVersion = item.ProductVersion ? item.ProductVersion : "V1";

          if (
            PrjVersion == currentProjectVersion &&
            PrdVersion == currentProductVersion
          ) {
            _activityPlanArr.push({
              ID: item.ID,
              Types: item.Types ? item.Types : "",
              Area: item.Area ? item.Area : "",
              ActivityPlanName: item.ActivityPlanName
                ? item.ActivityPlanName
                : "",
              Template: item.Title ? item.Title : "",
            });

            if (
              PLDrpDownOptns.Types.findIndex((_productDrpDwn) => {
                return _productDrpDwn.key == item.Types;
              }) == -1 &&
              item.Types
            ) {
              PLDrpDownOptns.Types.push({ key: item.Types, text: item.Types });
            }

            if (
              PLDrpDownOptns.Area.findIndex((_projectDrpDwn) => {
                return _projectDrpDwn.key == item.Area;
              }) == -1 &&
              item.Area
            ) {
              PLDrpDownOptns.Area.push({ key: item.Area, text: item.Area });
            }
          }
        });
        setPLAPDrpDownOptns({ ...PLDrpDownOptns });
        setPLAPMaster([..._activityPlanArr]);
        setPLAPData([..._activityPlanArr]);
        sortPLDataArr = [..._activityPlanArr];
        setPLAPFilData([..._activityPlanArr]);
        sortPLFilterArr = [..._activityPlanArr];
        paginate(1, [..._activityPlanArr]);
        setPLAPLoader("noLoader");
      })
      .catch((err) => {
        ErrorFunction(err, "getApData");
      });
  };

  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setPLAPDisplay(paginatedItems);
      setPLAPCurPage(pagenumber);
    } else {
      setPLAPDisplay([]);
      setPLAPCurPage(1);
    }
  };
  const PLFilterFunction = (key, value) => {
    let tempFilterKey = { ...PLAPFilterKey };
    tempFilterKey[key] = value;

    let tempArr = [...PLAPData];

    if (tempFilterKey.Types != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Types == tempFilterKey.Types;
      });
    }

    if (tempFilterKey.Area != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Area == tempFilterKey.Area;
      });
    }

    if (tempFilterKey.Template != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.Template.toLowerCase().includes(
          tempFilterKey.Template.toLowerCase()
        );
      });
    }
    if (tempFilterKey.ActivityPlanName != "") {
      tempArr = tempArr.filter((arr) => {
        return arr.ActivityPlanName != "";
      });
      tempArr = tempArr.filter((arr) => {
        return arr.ActivityPlanName.toLowerCase().includes(
          tempFilterKey.ActivityPlanName.toLowerCase()
        );
      });
    }
    setPLAPFilData([...tempArr]);
    sortPLFilterArr = [...tempArr];
    paginate(1, [...tempArr]);
    setPLAPFilterKey({ ...tempFilterKey });
  };
  const ErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    setPLAPLoader("noLoader");

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
    const tempapColumns = _PLAPColumns;
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
    setPLAPData([...newPLDataArr]);
    setPLAPFilData([...newPLFilterArr]);
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
    setPLAPLoader("startUpLoader");
    getActivityPlan();
  }, [PLAPReRender]);

  return (
    <>
      <div style={{ padding: "5px 15px" }}>
        {PLAPLoader == "startUpLoader" ? (
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
                  props.selectPLFunction("Product");
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
                  <Label style={{ fontSize: 24, padding: 0 }}>
                    Activity plan
                  </Label>
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
                  <Label className={PLlabelStyles.inputLabels}>Type</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      PLAPFilterKey.Types != "All"
                        ? PLActiveDropdownStyles
                        : PLDropdownStyles
                    }
                    options={PLAPDrpDownOptns.Types}
                    dropdownWidth={"auto"}
                    selectedKey={PLAPFilterKey.Types}
                    onChange={(e, option: any) => {
                      PLFilterFunction("Types", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>Area</Label>
                  <Dropdown
                    placeholder="Select an option"
                    styles={
                      PLAPFilterKey.Area != "All"
                        ? PLActiveDropdownStyles
                        : PLDropdownStyles
                    }
                    options={PLAPDrpDownOptns.Area}
                    dropdownWidth={"auto"}
                    selectedKey={PLAPFilterKey.Area}
                    onChange={(e, option: any) => {
                      PLFilterFunction("Area", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>
                    Activity plan name
                  </Label>
                  <SearchBox
                    styles={
                      PLAPFilterKey.ActivityPlanName
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLAPFilterKey.ActivityPlanName}
                    onChange={(e, value) => {
                      PLFilterFunction("ActivityPlanName", value);
                    }}
                  />
                </div>
                <div>
                  <Label className={PLlabelStyles.inputLabels}>Template</Label>
                  <SearchBox
                    styles={
                      PLAPFilterKey.Template
                        ? PLActiveSearchBoxStyles
                        : PLSearchBoxStyles
                    }
                    value={PLAPFilterKey.Template}
                    onChange={(e, value) => {
                      PLFilterFunction("Template", value);
                    }}
                  />
                </div>

                <div>
                  <Icon
                    iconName="Refresh"
                    title="Click to reset"
                    className={PLIconStyleClass.refresh}
                    onClick={() => {
                      setPLAPData([...PLAPMaster]);
                      sortPLDataArr = [...PLAPMaster];
                      setPLAPFilData([...PLAPMaster]);
                      sortPLFilterArr = [...PLAPMaster];
                      paginate(1, [...PLAPMaster]);
                      setPLAPFilterKey(PLFilterKey);
                      setPLAPColumns(_PLAPColumns);
                    }}
                  />
                </div>
              </div>
              <div>
                <Label className={PLlabelStyles.NORLabel}>
                  Number of records:{" "}
                  <b style={{ color: "#038387" }}>{PLAPFilData.length}</b>
                </Label>
              </div>
            </div>
            <div style={{ marginTop: "10px" }}>
              <DetailsList
                items={PLAPDisplay}
                columns={PLAPColumns}
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
                {PLAPFilData.length > 0 ? (
                  <Pagination
                    currentPage={PLAPcurPage}
                    totalPages={
                      PLAPFilData.length > 0
                        ? Math.ceil(PLAPFilData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginate(page, PLAPFilData);
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
export default PL_ActivityPlan;
