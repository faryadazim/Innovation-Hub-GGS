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
  Modal,
  PrimaryButton,
  TextField,
  ITextFieldStyles,
  Spinner,
  TooltipHost,
  TooltipOverflowMode,
  SearchBox,
  ISearchBoxStyles,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { arraysEqual, IDetailsListStyles } from "office-ui-fabric-react";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";

let columnSortArr = [];
let columnSortMasterArr = [];
let ProjectOrProductDetails = [];

const ActivityTemplate = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  let currentpage = 1;
  let totalPageItems = 10;

  let ATErrorStatus = {
    Types: "",
    Area: "",
    Product: "",
    Project: "",
    Lessons: "",
  };
  const ATDrpDwnOptns = {
    Project: [{ key: "All", text: "All" }],
    Types: [{ key: "All", text: "All" }],
    Area: [{ key: "All", text: "All" }],
    Product: [{ key: "All", text: "All" }],
  };
  const ATModalBoxDrpDwnOptns = {
    Project: [],
    Types: [],
    Area: [],
    Product: [],
  };
  const ATFilterKeys = {
    Project: "All",
    Types: "All",
    Area: "All",
    Product: "All",
    Code: "",
    Title: "",
  };
  const ATColumn = [
    {
      key: "Column1",
      name: "Type",
      fieldName: "Types",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Types}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Types}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column2",
      name: "Area/Stream",
      fieldName: "Area",
      minWidth: 150,
      maxWidth: 150,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Area}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Area}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column3",
      name: "Product(Program)",
      fieldName: "Product",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Product}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Product}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column4",
      name: "Project",
      fieldName: "Project",
      minWidth: 200,
      maxWidth: 400,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Project}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Project}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column5",
      name: "Code",
      fieldName: "Code",
      minWidth: 130,
      maxWidth: 130,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
    },
    {
      key: "Column6",
      name: "Template",
      fieldName: "Title",
      minWidth: 200,
      maxWidth: 300,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.Title}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.Title}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column7",
      name: "Action",
      fieldName: "",
      minWidth: 100,
      maxWidth: 150,
      onRender: (item) => (
        <>
          <Icon
            iconName="Edit"
            className={ATIconStyleClass.edit}
            onClick={() => {
              setATShowMessage(ATErrorStatus);
              let tempArr = {
                Id: item.Id,
                Title: item.Title,
                Project: item.Project,
                Types: item.Types,
                Area: item.Area,
                Product: item.Product,
                Code: item.Code,
                Lessons: item.Lessons,
              };
              if (item.Lessons) {
                let LessonArr = [];
                item.Lessons.split(";").forEach((le, index) => {
                  LessonArr.push({
                    Index: index,
                    Text: le,
                  });
                });
                setATLessons([...LessonArr]);
              } else {
                setATLessons([{ Index: 0, Text: "" }]);
              }

              setATModalBoxPopup({
                type: "Update",
                visible: true,
                value: tempArr,
              });
              // ATModalBoxDropDownOptions.Area.push({
              //   key: item.Area,
              //   text: item.Area,
              // });

              // Non dropdown value add
              let isOriginalData_Project = ProjectOrProductDetails.filter(
                (arr) => {
                  return (arr.Type = "Project" && arr.Key == item.Project);
                }
              );
              let isOriginalData_Product = ProjectOrProductDetails.filter(
                (arr) => {
                  return (arr.Type = "Product" && arr.Key == item.Product);
                }
              );

              if (isOriginalData_Product.length == 0) {
                ATModalBoxDropDownOptions.Product.push({
                  key: item.Product,
                  text: item.Product,
                });
              }
              if (isOriginalData_Project.length == 0) {
                ATModalBoxDropDownOptions.Project.push({
                  key: item.Project,
                  text: item.Project,
                });
              }

              setATModalBoxDropDownOptions({ ...ATModalBoxDropDownOptions });
            }}
          />
          <Icon
            iconName="Copy"
            className={ATIconStyleClass.edit}
            onClick={() => {
              setATShowMessage(ATErrorStatus);
              let tempArr = {
                Id: 0,
                Title: "",
                Project: item.Project,
                Types: item.Types,
                Area: item.Area,
                Product: item.Product,
                Code: item.Code,
                Lessons: item.Lessons,
              };
              if (item.Lessons) {
                let LessonArr = [];
                item.Lessons.split(";").forEach((le, index) => {
                  LessonArr.push({
                    Index: index,
                    Text: le,
                  });
                });
                setATLessons([...LessonArr]);
              } else {
                setATLessons([{ Index: 0, Text: "" }]);
              }

              setATModalBoxPopup({
                type: "New",
                visible: true,
                value: tempArr,
              });
              // ATModalBoxDropDownOptions.Area.push({
              //   key: item.Area,
              //   text: item.Area,
              // });

              // Non dropdown value add
              let isOriginalData_Project = ProjectOrProductDetails.filter(
                (arr) => {
                  return (arr.Type = "Project" && arr.Key == item.Project);
                }
              );
              let isOriginalData_Product = ProjectOrProductDetails.filter(
                (arr) => {
                  return (arr.Type = "Product" && arr.Key == item.Product);
                }
              );

              if (isOriginalData_Product.length == 0) {
                ATModalBoxDropDownOptions.Product.push({
                  key: item.Product,
                  text: item.Product,
                });
              }
              if (isOriginalData_Project.length == 0) {
                ATModalBoxDropDownOptions.Project.push({
                  key: item.Project,
                  text: item.Project,
                });
              }
              setATModalBoxDropDownOptions({ ...ATModalBoxDropDownOptions });
            }}
          />
          <Icon
            iconName="Delete"
            className={ATIconStyleClass.delete}
            onClick={() => {
              setATDeletePopup({ condition: true, targetId: item.Id });
              console.log(ATDeletePopup);
            }}
          />
        </>
      ),
    },
  ];
  const ATIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const ATIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, ATIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, ATIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, ATIconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 29,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
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
  const ATProjectInfo = mergeStyles({
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 10px",
  });
  // detailslist
  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          ".ms-DetailsRow-fields": {
            alignItems: "center",
          },
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };
  const ATBigiconStyleClass = mergeStyleSets({
    ChevronLeftMed: [
      {
        cursor: "pointer",
        color: "#2392b2",
        fontSize: 24,
        marginTop: "3px",
        marginRight: 12,
      },
    ],
  });
  const ATbuttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
  });
  const ATbuttonStyleClass = mergeStyleSets({
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
      ATbuttonStyle,
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
      ATbuttonStyle,
    ],
  });
  const ATSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const ATActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
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
  const ATlabelStyles = mergeStyleSets({
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
        // marginTop: "25px",
        marginLeft: "10px",
        fontWeight: "500",
        color: "#323130",
        fontSize: "13px",
      },
    ],
  });
  const ATdropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      border: "1px solid #E8E8EA",
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ATActivedropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      fontWeight: 600,
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };
  const ATReadOnlydropdownStyles: Partial<IDropdownStyles> = {
    root: { width: 186, marginRight: 15 },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "2px solid #7C7C7C",
      fontWeight: 600,
    },
    dropdownItemsWrapper: { backgroundColor: "#F5F5F7", fontSize: 12 },
    dropdownItemSelected: { backgroundColor: "#DCDCDC", fontSize: 12 },
    caretDown: { fontSize: 14, color: "#7C7C7C", display: "none" },
  };
  const ATModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      padding: "3px 10px",
      height: "36px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#000",
      fontWeight: "bold",
    },
    callout: { height: 200 },
  };
  const ATModalBoxActiveDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      padding: "3px 10px",
      height: "36px",
      color: "#038387",
      border: "2px solid #038387",
      fontWeight: "bold",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#038387",
      fontWeight: "bold",
    },
    callout: { maxHeight: 200 },
  };
  const ATModalBoxReadOnlyDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      padding: "3px 10px",
      height: "36px",
      color: "#7C7C7C",
      border: "2px solid #7C7C7C",
      fontWeight: "bold",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#7C7C7C",
      fontWeight: "bold",
      display: "none",
    },
    callout: { maxHeight: 200 },
  };
  const ATTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: 200, marginLeft: 5 },
    field: { backgroundColor: "whitesmoke", fontSize: 12 },
  };

  // Functions
  const getActivityTemplateList = () => {
    let _ATdata = [];
    sharepointWeb.lists
      .getByTitle("Activity Template")
      .items.filter("IsDeleted ne 1")
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          _ATdata.push({
            Id: item.Id,
            Title: item.Title,
            Project:
              item.Project +
              " " +
              (item.ProjectVersion ? item.ProjectVersion : "V1"),
            Types: item.Types,
            Area: item.Area,
            Product:
              item.Product +
              " " +
              (item.ProductVersion ? item.ProductVersion : "V1"),
            Code: item.Code,
            Lessons: item.Lessons,
          });
        });
        setATUnsortMasterData([..._ATdata]);
        columnSortArr = _ATdata;
        setATData([..._ATdata]);
        columnSortMasterArr = _ATdata;
        setATMasterData([..._ATdata]);
        reloadFilterOptions([..._ATdata]);
        paginate(1, [..._ATdata]);
        setATLoader(false);
      })
      .catch((err) => {
        ATErrorFunction(err, "getActivityTemplateList");
      });
  };
  const getProductList = () => {
    let _PLdata = [];
    // let tempOption = ATModalBoxDrpDwnOptns;
    sharepointWeb.lists
      .getByTitle("Product List")
      .items.top(5000)
      .get()
      .then(async (items) => {
        console.log(items);
        items.forEach((item) => {
          _PLdata.push({
            Types: item.Types,
            Area: item.Title,
            Product: item.Product,
            Project: item.Project,
            Code: item.Code,
            Publisher: item.Publisher,
            Status: item.Status,
          });

          if (
            ATModalBoxDrpDwnOptns.Area.findIndex((area) => {
              return area.key == item.Title;
            }) == -1 &&
            item.Title
          ) {
            ATModalBoxDrpDwnOptns.Area.push({
              key: item.Title,
              text: item.Title,
            });
          }
        });

        setProductData([..._PLdata]);
        getTemplateOptions();

        // setATModalBoxDropDownOptions(tempOption);
        // getActivityTemplateList();
      })
      .catch((err) => {
        ATErrorFunction(err, "getProductList");
      });
  };
  const getTemplateOptions = () => {
    ProjectOrProductDetails = [];
    const _sortFilterKeys = (a, b) => {
      if (a.text.toLowerCase() < b.text.toLowerCase()) {
        return -1;
      }
      if (a.text.toLowerCase() > b.text.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    //Types Choices
    sharepointWeb.lists
      .getByTitle("Product List")
      .fields.getByInternalNameOrTitle("Types")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ATModalBoxDrpDwnOptns.Types.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ATModalBoxDrpDwnOptns.Types.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {})
      .catch((err) => {
        ATErrorFunction(err, "getTemplateOptions-Types");
      });

    //Product Choices
    sharepointWeb.lists
      .getByTitle("Master Product List")
      .items.filter("IsDeleted ne 1")
      .top(5000)
      .get()
      .then((allProducts) => {
        allProducts.forEach((product) => {
          if (product.Title != null) {
            if (
              ATModalBoxDrpDwnOptns.Product.findIndex((productOptn) => {
                return productOptn.key == product.Title;
              }) == -1
            ) {
              if (product.Title != "Not Sure") {
                ATModalBoxDrpDwnOptns.Product.push({
                  key: product.Title + " " + product.ProductVersion,
                  text: product.Title + " " + product.ProductVersion,
                });
              }
              ProjectOrProductDetails.push({
                Type: "Product",
                Id: product.ID,
                Key: product.Title + " " + product.ProductVersion,
                Title: product.Title,
                Version: product.ProductVersion,
              });
            }
          }
        });
      })
      .then(() => {
        ATModalBoxDrpDwnOptns.Product.sort(_sortFilterKeys);
        ATModalBoxDrpDwnOptns.Product.unshift({
          key: "Not Sure V1",
          text: "Not Sure V1",
        });
      })
      .catch((err) => {
        ATErrorFunction(err, "getTemplateOptions-Product");
      });

    //Project & Product Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.top(5000)
      .orderBy("Modified", false)
      .get()
      .then((Items) => {
        Items.forEach((arr) => {
          if (
            ATModalBoxDrpDwnOptns.Project.findIndex((prj) => {
              return prj.key == arr.Title;
            }) == -1 &&
            arr.Title
          ) {
            ATModalBoxDrpDwnOptns.Project.push({
              key: arr.Title + " " + arr.ProjectVersion,
              text: arr.Title + " " + arr.ProjectVersion,
            });
            ProjectOrProductDetails.push({
              Type: "Project",
              Id: arr.ID,
              Key: arr.Title + " " + arr.ProjectVersion,
              Title: arr.Title,
              Version: arr.ProjectVersion,
            });
          }
        });
      })
      .then(() => {
        ATModalBoxDrpDwnOptns.Project.sort(_sortFilterKeys);
      })
      .catch((err) => {
        ATErrorFunction(err, "getTemplateOptions-Project");
      });
    setATModalBoxDropDownOptions(ATModalBoxDrpDwnOptns);
    // tempArr.forEach((arr) => {
    //   if (
    //     ATModalBoxDrpDwnOptns.Area.findIndex((area) => {
    //       return area.key == arr.Area;
    //     }) == -1 &&
    //     arr.Area
    //   ) {
    //     ATModalBoxDrpDwnOptns.Area.push({
    //       key: arr.Area,
    //       text: arr.Area,
    //     });
    //   }
    //   if (
    //     ATModalBoxDrpDwnOptns.Product.findIndex((prd) => {
    //       return prd.key == arr.Product;
    //     }) == -1 &&
    //     arr.Product
    //   ) {
    //     ATModalBoxDrpDwnOptns.Product.push({
    //       key: arr.Product,
    //       text: arr.Product,
    //     });
    //   }
    //   if (
    //     ATModalBoxDrpDwnOptns.Project.findIndex((prj) => {
    //       return prj.key == arr.Project;
    //     }) == -1 &&
    //     arr.Project
    //   ) {
    //     ATModalBoxDrpDwnOptns.Project.push({
    //       key: arr.Project,
    //       text: arr.Project,
    //     });
    //   }
    // });
    // setATModalBoxDropDownOptions(ATModalBoxDrpDwnOptns);
  };
  const reloadFilterOptions = (data) => {
    let tempArrReload = data;

    tempArrReload.forEach((at) => {
      if (
        at.Project &&
        ATDrpDwnOptns.Project.findIndex((prj) => {
          return prj.key == at.Project;
        }) == -1
      ) {
        ATDrpDwnOptns.Project.push({
          key: at.Project,
          text: at.Project,
        });
      }
      if (
        at.Types &&
        ATDrpDwnOptns.Types.findIndex((type) => {
          return type.key == at.Types;
        }) == -1
      ) {
        ATDrpDwnOptns.Types.push({
          key: at.Types,
          text: at.Types,
        });
      }
      if (
        at.Area &&
        ATDrpDwnOptns.Area.findIndex((area) => {
          return area.key == at.Area;
        }) == -1
      ) {
        ATDrpDwnOptns.Area.push({
          key: at.Area,
          text: at.Area,
        });
      }
      if (
        at.Product &&
        ATDrpDwnOptns.Product.findIndex((prd) => {
          return prd.key == at.Product;
        }) == -1
      ) {
        ATDrpDwnOptns.Product.push({
          key: at.Product,
          text: at.Product,
        });
      }
    });
    setATDropDownOptions(ATDrpDwnOptns);
  };
  const ATDeleteItem = (id: number) => {
    // sharepointWeb.lists
    //   .getByTitle("Activity Template")
    //   .items.getById(id)
    //   .delete()
    //   .then(() => {
    sharepointWeb.lists
      .getByTitle("Activity Template")
      .items.getById(id)
      .update({
        IsDeleted: true,
      })
      .then((e) => {
        let tempMasterArr = [...ATMasterData];
        let targetIndex = tempMasterArr.findIndex((arr) => arr.Id == id);
        tempMasterArr.splice(targetIndex, 1);

        let tempdataarr = [...ATData];
        let targetIndexatdata = tempdataarr.findIndex((arr) => arr.Id == id);
        tempdataarr.splice(targetIndexatdata, 1);
        setATUnsortMasterData([...tempMasterArr]);
        columnSortArr = tempMasterArr;
        setATData([...tempMasterArr]);
        columnSortMasterArr = tempMasterArr;
        setATMasterData([...tempMasterArr]);
        paginate(1, [...tempdataarr]);
        reloadFilterOptions([...tempdataarr]);
        setATDeletePopup({ condition: false, targetId: 0 });
        setATOnSubmitLoader(false);
        DeleteSuccessPopup();
      })
      .catch((err) => {
        ATErrorFunction(err, "ATDeleteItem");
      });
  };
  const sortFilterKeys = (b, a) => {
    if (a.Id < b.Id) {
      return -1;
    }
    if (a.Id > b.Id) {
      return 1;
    }
    return 0;
  };
  const ATSaveData = () => {
    let itemID: number = ATModalBoxPopup.value["Id"];

    //Template TypesCount
    let typesCount = "01";
    let CurType = ATMasterData.filter((at) => {
      return at.Types == ATModalBoxPopup.value["Types"];
    });
    if (CurType.length > 0) {
      if (CurType.length > 1) {
        CurType = ATMasterData.filter((at) => {
          return at.Types == ATModalBoxPopup.value["Types"];
        }).sort(sortFilterKeys);

        let CurTitle = CurType[0].Title.split("-")[1];
        typesCount =
          parseInt(CurTitle) < 9
            ? "0" + (parseInt(CurTitle) + 1).toString()
            : (parseInt(CurTitle) + 1).toString();
      } else {
        let CurTitle = CurType[0].Title.split("-")[1];
        typesCount =
          parseInt(CurTitle) < 9
            ? "0" + (parseInt(CurTitle) + 1).toString()
            : (parseInt(CurTitle) + 1).toString();
      }
    }

    debugger;

    // Template Code
    let codeValueTypes = ProductData.filter((pd) => {
      return pd.Types.find((type) => {
        return type == ATModalBoxPopup.value["Types"];
      });
    });
    let codeValue = codeValueTypes.filter((pd) => {
      return (
        //pd.Types == ATModalBoxPopup.value["Types"] &&
        pd.Area == ATModalBoxPopup.value["Area"] &&
        pd.Product == ATModalBoxPopup.value["Product"] &&
        pd.Project == ATModalBoxPopup.value["Project"]
      );
    });

    // Versions
    let PrjData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type =
        "Project" && arr.Key == ATModalBoxPopup.value["Project"]);
    });
    let PrdData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type =
        "Product" && arr.Key == ATModalBoxPopup.value["Product"]);
    });

    let PrjTitle =
      PrjData.length > 0
        ? PrjData[0].Title
        : ATModalBoxPopup.value["Project"].replace("V1", "");
    let PrjVersion = PrjData.length > 0 ? PrjData[0].Version : "V1";

    let PrdTitle =
      PrdData.length > 0
        ? PrdData[0].Title
        : ATModalBoxPopup.value["Product"].replace("V1", "");
    let PrdVersion = PrdData.length > 0 ? PrdData[0].Version : "V1";

    let requestdata = {
      Title: ATModalBoxPopup.value["Title"]
        ? ATModalBoxPopup.value["Title"]
        : ATModalBoxPopup.value["Types"] + " - " + typesCount,
      Types: ATModalBoxPopup.value["Types"]
        ? ATModalBoxPopup.value["Types"]
        : null,
      Area: ATModalBoxPopup.value["Area"]
        ? ATModalBoxPopup.value["Area"]
        : null,
      Product: PrdTitle,
      Project: PrjTitle,
      ProductVersion: PrdVersion,
      ProjectVersion: PrjVersion,
      Code: codeValue.length > 0 ? codeValue[0].Code : null,
      Lessons: ATModalBoxPopup.value["Lessons"]
        ? ATModalBoxPopup.value["Lessons"]
        : null,
    };

    ATModalBoxPopup.type == "New"
      ? sharepointWeb.lists
          .getByTitle("Activity Template")
          .items.add(requestdata)
          .then((e) => {
            setATModalBoxPopup({
              type: "",
              visible: false,
              value: {},
            });
            ATModalBoxPopup.value["Code"] =
              codeValue.length > 0 ? codeValue[0].Code : null;
            ATModalBoxPopup.value["Title"] = ATModalBoxPopup.value["Title"]
              ? ATModalBoxPopup.value["Title"]
              : ATModalBoxPopup.value["Types"] + " - " + typesCount;
            ATModalBoxPopup.value["Id"] = e.data.ID;

            ATData.unshift(ATModalBoxPopup.value);
            ATMasterData.unshift(ATModalBoxPopup.value);

            setATUnsortMasterData([...ATMasterData]);
            columnSortArr = ATData;
            setATData([...ATData]);
            columnSortArr = ATMasterData;
            setATMasterData([...ATMasterData]);
            paginate(1, [...ATData]);
            reloadFilterOptions([...ATData]);
            setATOnSubmitLoader(false);
            AddSuccessPopup();
          })
          .catch((err) => {
            ATErrorFunction(err, "ATSaveData-add");
          })
      : sharepointWeb.lists
          .getByTitle("Activity Template")
          .items.getById(itemID)
          .update(requestdata)
          .then((e) => {
            setATModalBoxPopup({
              type: "",
              visible: false,
              value: {},
            });
            let disIndex = ATData.findIndex((obj) => obj.Id == itemID);
            let Index = ATMasterData.findIndex((obj) => obj.Id == itemID);

            ATData.splice(disIndex, 1);
            ATMasterData.splice(Index, 1);

            ATData.unshift(ATModalBoxPopup.value);
            ATMasterData.unshift(ATModalBoxPopup.value);

            setATUnsortMasterData([...ATMasterData]);
            columnSortArr = ATData;
            setATData([...columnSortArr]);
            columnSortMasterArr = ATMasterData;
            setATMasterData([...ATMasterData]);
            paginate(1, [...ATData]);
            reloadFilterOptions([...ATData]);
            setATOnSubmitLoader(false);
            UpdateSuccessPopup();
          })
          .catch((err) => {
            ATErrorFunction(err, "ATSaveData-update");
          });
  };
  const paginate = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setATDisplayData(paginatedItems);
      setATCurrentPage(pagenumber);
    } else {
      setATDisplayData([]);
      setATCurrentPage(1);
    }
  };
  const ATValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      Types: "",
      Area: "",
      Product: "",
      Project: "",
      Lessons: "",
    };

    if (!ATModalBoxPopup.value["Types"]) {
      isError = true;
      errorStatus.Types = "Please select a value for Type";
    }
    // if (!ATModalBoxPopup.value["Area"]) {
    //   isError = true;
    //   errorStatus.Area = "Please select a value for Area";
    // }
    if (!ATModalBoxPopup.value["Product"]) {
      isError = true;
      errorStatus.Product = "Please select a value for Product";
    }
    if (!ATModalBoxPopup.value["Project"]) {
      isError = true;
      errorStatus.Project = "Please select a value for Project";
    }
    if (!ATModalBoxPopup.value["Lessons"]) {
      isError = true;
      errorStatus.Lessons = "Please add section";
    }
    if (!isError) {
      setATOnSubmitLoader(true);
      setATShowMessage(ATErrorStatus);

      // Non dropdown value remove
      // Non dropdown value remove
      let isOriginalData_Project = ProjectOrProductDetails.filter((arr) => {
        return (arr.Type =
          "Project" && arr.Key == ATModalBoxPopup["value"]["Project"]);
      });
      let isOriginalData_Product = ProjectOrProductDetails.filter((arr) => {
        return (arr.Type =
          "Product" && arr.Key == ATModalBoxPopup["value"]["Product"]);
      });
      if (
        isOriginalData_Project.length == 0 &&
        ATModalBoxPopup["value"]["Project"]
      ) {
        ATModalBoxDropDownOptions.Project.pop();
      }
      if (
        isOriginalData_Product.length == 0 &&
        ATModalBoxPopup["value"]["Product"]
      ) {
        ATModalBoxDropDownOptions.Product.pop();
      }

      setATModalBoxDropDownOptions({
        ...ATModalBoxDropDownOptions,
      });

      ATSaveData();
    } else {
      setATOnSubmitLoader(false);
      setATShowMessage(errorStatus);
    }
  };
  const ATErrorFunction = (error, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Activity Template",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setATOnSubmitLoader(false);
        setATLoader(false);
        ErrorPopup();
      }
    );
  };
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity template is successfully submitted !!!")
  );
  const UpdateSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity template is successfully updated !!!")
  );
  const DeleteSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity template is successfully deleted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  // Onchange and Filters
  const ATListFilter = (key, option) => {
    let tempArr = [...ATMasterData];
    let tempDpFilterKeys = { ...ATFilterOptions };
    tempDpFilterKeys[`${key}`] = option;

    if (tempDpFilterKeys.Project != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Project == tempDpFilterKeys.Project;
      });
    }
    if (tempDpFilterKeys.Types != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Types == tempDpFilterKeys.Types;
      });
    }
    if (tempDpFilterKeys.Area != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Area == tempDpFilterKeys.Area;
      });
    }
    if (tempDpFilterKeys.Product != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Product == tempDpFilterKeys.Product;
      });
    }
    if (tempDpFilterKeys.Title) {
      tempArr = tempArr.filter((arr) => {
        return arr.Title.toLowerCase().includes(
          tempDpFilterKeys.Title.toLowerCase()
        );
      });
    }
    if (tempDpFilterKeys.Code) {
      tempArr = tempArr.filter((arr) => {
        if (arr.Code) {
          return arr.Code.toLowerCase().includes(
            tempDpFilterKeys.Code.toLowerCase()
          );
        }
      });
    }

    columnSortArr = tempArr;
    setATData([...columnSortArr]);
    paginate(1, [...tempArr]);
    setATFilterOptions({ ...tempDpFilterKeys });
  };
  const ATOnchange = (key, value) => {
    let tempArr = {
      Id: ATModalBoxPopup.value["Id"],
      Title: ATModalBoxPopup.value["Title"],
      Project: key == "Project" ? value : ATModalBoxPopup.value["Project"],
      Types: key == "Types" ? value : ATModalBoxPopup.value["Types"],
      Area: key == "Area" ? value : ATModalBoxPopup.value["Area"],
      Product: key == "Product" ? value : ATModalBoxPopup.value["Product"],
      Code: ATModalBoxPopup.value["Code"],
      Lessons: key == "Lessons" ? value : ATModalBoxPopup.value["Lessons"],
    };
    setATModalBoxPopup({
      type: ATModalBoxPopup.type,
      visible: true,
      value: tempArr,
    });
    console.log(tempArr);
    // ATCascadingFilter(key, value);
  };
  const ATCascadingFilter = (key, value) => {
    if (key == "Types") {
      let tempMulArr = ProductData.filter((arr) => {
        return arr[key].find((type) => {
          return type == value;
        });
      });

      ATModalBoxDropDownOptions.Area = [];
      ATModalBoxDropDownOptions.Product = [];
      ATModalBoxDropDownOptions.Project = [];

      tempMulArr.forEach((arr) => {
        if (
          ATModalBoxDropDownOptions.Area.findIndex((area) => {
            return area.key == arr.Area;
          }) == -1 &&
          arr.Area
        ) {
          ATModalBoxDropDownOptions.Area.push({
            key: arr.Area,
            text: arr.Area,
          });
        }
      });
    } else {
      let tempArr = ProductData.filter((arr) => {
        return arr[key] == value;
      });

      if (key == "Area") {
        ATModalBoxDropDownOptions.Product = [];
        ATModalBoxDropDownOptions.Project = [];
        tempArr.forEach((arr) => {
          if (
            ATModalBoxDropDownOptions.Product.findIndex((prd) => {
              return prd.key == arr.Product;
            }) == -1 &&
            arr.Product
          ) {
            ATModalBoxDropDownOptions.Product.push({
              key: arr.Product,
              text: arr.Product,
            });
          }
        });
      } else if (key == "Product") {
        ATModalBoxDropDownOptions.Project = [];
        tempArr.forEach((arr) => {
          if (
            ATModalBoxDropDownOptions.Project.findIndex((prj) => {
              return prj.key == arr.Project;
            }) == -1 &&
            arr.Project
          ) {
            ATModalBoxDropDownOptions.Project.push({
              key: arr.Project,
              text: arr.Project,
            });
          }
        });
      } else if (key == "Project" || key == "Lessons") {
      } else {
        ATModalBoxDropDownOptions.Area = [];
        ATModalBoxDropDownOptions.Product = [];
        ATModalBoxDropDownOptions.Project = [];
      }
    }
    setATModalBoxDropDownOptions(ATModalBoxDropDownOptions);
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = ATColumn;
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

    const newDisplayData = _copyAndSort(
      columnSortArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newMasterData = _copyAndSort(
      columnSortMasterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setATData([...newDisplayData]);
    setATMasterData([...newMasterData]);
    paginate(1, newDisplayData);
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

  //Use State
  const [ATReRender, setATReRender] = useState(false);
  const [ATData, setATData] = useState([]);
  const [ATUnsortMasterData, setATUnsortMasterData] = useState([]);
  const [ATMasterData, setATMasterData] = useState([]);
  const [ATDisplayData, setATDisplayData] = useState([]);
  const [ATcurrentPage, setATCurrentPage] = useState(currentpage);
  const [ProductData, setProductData] = useState([]);
  const [ATDropDownOptions, setATDropDownOptions] = useState(ATDrpDwnOptns);
  const [ATFilterOptions, setATFilterOptions] = useState(ATFilterKeys);
  const [ATDeletePopup, setATDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });
  const [ATLoader, setATLoader] = useState(true);
  const [ATOnSubmitLoader, setATOnSubmitLoader] = useState(false);
  const [ATModalBoxPopup, setATModalBoxPopup] = useState({
    type: "",
    visible: false,
    value: {},
  });
  const [ATModalBoxDropDownOptions, setATModalBoxDropDownOptions] = useState(
    ATModalBoxDrpDwnOptns
  );
  const [ATLessons, setATLessons] = useState([]);
  const [ATPopup, setATPopup] = useState("");
  const [ATShowMessage, setATShowMessage] = useState(ATErrorStatus);
  const [ATMasterColumns, setATMasterColumns] = useState(ATColumn);

  // Use Effect
  useEffect(() => {
    // console.log(ATLessons.length);
    getProductList();
    getActivityTemplateList();
  }, [ATReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {ATLoader ? <CustomLoader /> : null}
      <div className={styles.attHeaderSection}>
        <div
          style={{
            display: "flex",
            alignItems: "flex-start",
            justifyContent: "space-between",
            marginBottom: 10,
            color: "#2392b2",
          }}
        >
          {/* Header Start */}
          <div className={styles.dpTitle}>
            <Icon
              aria-label="ChevronLeftMed"
              iconName="NavigateBack"
              className={ATBigiconStyleClass.ChevronLeftMed}
              onClick={() => {
                props.handleclick("ActivityPlan", null, "ATP");
              }}
            />
            <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
              Activity template
            </Label>
          </div>
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "left",
            marginBottom: "5px",
          }}
        >
          <PrimaryButton
            text="Add"
            className={ATbuttonStyleClass.buttonPrimary}
            onClick={(_) => {
              setATShowMessage(ATErrorStatus);
              setATModalBoxPopup({
                type: "New",
                visible: true,
                value: {},
              });
              // ATCascadingFilter(null, null);
              setATLessons([{ Index: 0, Text: "" }]);
            }}
          />
        </div>
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            flexWrap: "wrap",
          }}
        >
          <div className={styles.ddSection}>
            <div>
              <Label className={ATlabelStyles.inputLabels}>Type</Label>
              <Dropdown
                selectedKey={ATFilterOptions.Types}
                placeholder="Select an option"
                options={ATDropDownOptions.Types}
                styles={ATdropdownStyles}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  ATListFilter("Types", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={ATlabelStyles.inputLabels}>Area/Stream</Label>
              <Dropdown
                selectedKey={ATFilterOptions.Area}
                placeholder="Select an option"
                options={ATDropDownOptions.Area}
                styles={ATdropdownStyles}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  ATListFilter("Area", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={ATlabelStyles.inputLabels}>
                Product(Program)
              </Label>
              <Dropdown
                selectedKey={ATFilterOptions.Product}
                placeholder="Select an option"
                options={ATDropDownOptions.Product}
                styles={ATdropdownStyles}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  ATListFilter("Product", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={ATlabelStyles.inputLabels}>Project</Label>
              <Dropdown
                selectedKey={ATFilterOptions.Project}
                placeholder="Select an option"
                options={ATDropDownOptions.Project}
                styles={ATdropdownStyles}
                dropdownWidth={"auto"}
                onChange={(e, option: any) => {
                  ATListFilter("Project", option["key"]);
                }}
              />
            </div>
            <div>
              <Label className={ATlabelStyles.inputLabels}>Code</Label>
              <SearchBox
                styles={
                  ATFilterOptions.Code
                    ? ATActiveSearchBoxStyles
                    : ATSearchBoxStyles
                }
                value={ATFilterOptions.Code}
                onChange={(e, value) => {
                  ATListFilter("Code", value);
                }}
              />
            </div>
            <div>
              <Label className={ATlabelStyles.inputLabels}>Template</Label>
              <SearchBox
                styles={
                  ATFilterOptions.Title
                    ? ATActiveSearchBoxStyles
                    : ATSearchBoxStyles
                }
                value={ATFilterOptions.Title}
                onChange={(e, value) => {
                  ATListFilter("Title", value);
                }}
              />
            </div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={ATIconStyleClass.refresh}
                onClick={() => {
                  setATFilterOptions(ATFilterKeys);
                  columnSortArr = ATUnsortMasterData;
                  setATData([...ATUnsortMasterData]);
                  columnSortMasterArr = ATUnsortMasterData;
                  setATMasterData([...ATUnsortMasterData]);
                  paginate(1, [...ATUnsortMasterData]);
                  setATMasterColumns(ATColumn);
                }}
              />
            </div>
          </div>
          <div
            className={ATProjectInfo}
            style={{
              marginLeft: "20px",
              transform: "translateY(12px)",
            }}
          >
            <Label className={ATlabelStyles.NORLabel}>
              Number of records:{" "}
              <b style={{ color: "#038387" }}>{ATData.length}</b>
            </Label>
          </div>
        </div>
      </div>
      <DetailsList
        items={ATDisplayData}
        columns={ATMasterColumns}
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
        {ATData.length > 0 ? (
          <Pagination
            currentPage={ATcurrentPage}
            totalPages={
              ATData.length > 0 ? Math.ceil(ATData.length / totalPageItems) : 1
            }
            onChange={(page) => {
              paginate(page, ATData);
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
            <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
          </div>
        )}
      </div>
      {/* Delete Popup */}
      <div>
        {ATDeletePopup.condition ? (
          <Modal isOpen={ATDeletePopup.condition} isBlocking={true}>
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
                  justifyContent: "flex-start",
                  flexDirection: "column",
                  marginBottom: "10px",
                }}
              >
                <Label className={styles.deletePopupTitle}>
                  Delete template
                </Label>
                <Label className={styles.deletePopupDesc}>
                  Are you sure you want to delete this template?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setATOnSubmitLoader(true);
                  ATDeleteItem(ATDeletePopup.targetId);
                }}
                className={styles.apDeletePopupYesBtn}
              >
                {ATOnSubmitLoader ? <Spinner /> : "Yes"}
              </button>
              <button
                onClick={(_) => {
                  setATDeletePopup({ condition: false, targetId: 0 });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        ) : (
          ""
        )}
      </div>
      <div>
        <Modal isOpen={ATModalBoxPopup.visible} isBlocking={false}>
          <div style={{ padding: "30px 20px" }}>
            <div
              style={{
                fontSize: 24,
                textAlign: "center",
                color: "#2392B2",
                fontWeight: "600",
                marginBottom: "20px",
              }}
            >
              {ATModalBoxPopup.type == "New"
                ? "New template"
                : "Update template"}
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "flex-start",
              }}
            >
              <div>
                <Dropdown
                  required={true}
                  errorMessage={ATShowMessage.Types}
                  label="Type"
                  // dropdownWidth={"auto"}
                  selectedKey={ATModalBoxPopup.value["Types"]}
                  placeholder="Select an option"
                  options={ATModalBoxDropDownOptions.Types}
                  styles={
                    // ATModalBoxPopup.value["Types"] &&
                    // ATModalBoxPopup.type == "Update"
                    //   ? ATModalBoxReadOnlyDrpDwnCalloutStyles
                    //   : ATModalBoxPopup.value["Types"]
                    //   ? ATModalBoxActiveDrpDwnCalloutStyles
                    //   : ATModalBoxDrpDwnCalloutStyles
                    ATModalBoxDrpDwnCalloutStyles
                  }
                  disabled={ATModalBoxPopup.type == "New" ? false : true}
                  onChange={(e, option: any) => {
                    ATOnchange("Types", option["key"]);
                  }}
                />
              </div>
              <div>
                <Dropdown
                  // required={true}
                  // errorMessage={ATShowMessage.Area}
                  label="Area/Stream"
                  // dropdownWidth={"auto"}
                  selectedKey={ATModalBoxPopup.value["Area"]}
                  placeholder="Select an option"
                  options={ATModalBoxDropDownOptions.Area}
                  styles={
                    // ATModalBoxPopup.value["Area"] &&
                    // ATModalBoxPopup.type == "Update"
                    //   ? ATModalBoxReadOnlyDrpDwnCalloutStyles
                    //   : ATModalBoxPopup.value["Area"]
                    //   ? ATModalBoxActiveDrpDwnCalloutStyles
                    //   : ATModalBoxDrpDwnCalloutStyles
                    ATModalBoxDrpDwnCalloutStyles
                  }
                  disabled={ATModalBoxPopup.type == "New" ? false : true}
                  onChange={(e, option: any) => {
                    ATOnchange("Area", option["key"]);
                  }}
                />
              </div>
              <div>
                <Dropdown
                  required={true}
                  errorMessage={ATShowMessage.Product}
                  label="Product(Program)"
                  // dropdownWidth={"auto"}
                  selectedKey={ATModalBoxPopup.value["Product"]}
                  placeholder="Select an option"
                  options={ATModalBoxDropDownOptions.Product}
                  styles={
                    // ATModalBoxPopup.value["Product"] &&
                    // ATModalBoxPopup.type == "Update"
                    //   ? ATModalBoxReadOnlyDrpDwnCalloutStyles
                    //   : ATModalBoxPopup.value["Product"]
                    //   ? ATModalBoxActiveDrpDwnCalloutStyles
                    //   : ATModalBoxDrpDwnCalloutStyles
                    ATModalBoxDrpDwnCalloutStyles
                  }
                  disabled={ATModalBoxPopup.type == "New" ? false : true}
                  onChange={(e, option: any) => {
                    ATOnchange("Product", option["key"]);
                  }}
                />
              </div>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "flex-start",
              }}
            >
              <div>
                <Dropdown
                  required={true}
                  errorMessage={ATShowMessage.Project}
                  label="Project"
                  // dropdownWidth={"auto"}
                  selectedKey={ATModalBoxPopup.value["Project"]}
                  placeholder="Select an option"
                  options={ATModalBoxDropDownOptions.Project}
                  styles={
                    // ATModalBoxPopup.value["Project"] &&
                    // ATModalBoxPopup.type == "Update"
                    //   ? ATModalBoxReadOnlyDrpDwnCalloutStyles
                    //   : ATModalBoxPopup.value["Project"]
                    //   ? ATModalBoxActiveDrpDwnCalloutStyles
                    //   : ATModalBoxDrpDwnCalloutStyles
                    ATModalBoxDrpDwnCalloutStyles
                  }
                  disabled={ATModalBoxPopup.type == "New" ? false : true}
                  style={{ width: "640px" }}
                  onChange={(e, option: any) => {
                    ATOnchange("Project", option["key"]);
                  }}
                />
              </div>
            </div>

            <Label
              style={{
                marginLeft: 21,
              }}
              required={true}
            >
              Section
            </Label>

            <div
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "flex-start",
                minWidth: "300px",
                maxWidth: "1000px",
                height: "130px",
                flexWrap: "wrap",
                overflow: "auto",
              }}
              className={styles.addSection}
            >
              {ATLessons.length > 0
                ? ATLessons.map((les, actualIndex) => {
                    return (
                      <div
                        style={{
                          display: "flex",
                          alignItems: "flex-start",
                          justifyContent: "flex-start",
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                            backgroundColor: "#eeeeee80",
                            padding: "10px",
                            margin: "5px",
                            marginRight: "20px",
                            borderRadius: "4px",
                            transform: "translateX(20px)",
                          }}
                        >
                          <TextField
                            styles={ATTxtBoxStyles}
                            errorMessage={
                              actualIndex == 0 ? ATShowMessage.Lessons : null
                            }
                            value={les.Text}
                            data-index={les.Index}
                            onChange={(e, value: string) => {
                              let dataindex =
                                e.target["getAttribute"]("data-index");
                              let Index = ATLessons.findIndex(
                                (le) => le.Index == dataindex
                              );
                              ATLessons[Index].Text = value;
                              setATLessons([...ATLessons]);
                              ATOnchange(
                                "Lessons",
                                ATLessons.map((user) => user.Text).join(";")
                              );
                            }}
                          />
                          <div
                            style={{
                              display: "flex",
                              alignItems: "center",
                              justifyContent: "center",
                              marginLeft: "20px",
                            }}
                          >
                            <Icon
                              aria-label="CircleAddition"
                              iconName="CircleAddition"
                              data-index={les.Index}
                              className={ATIconStyleClass.link}
                              onClick={() => {
                                let addItems = [...ATLessons];
                                addItems.push({
                                  // Index: les.Index + 1,
                                  Index: ATLessons.length,
                                  Text: "",
                                });
                                setATLessons([...addItems]);
                                console.log(addItems);
                              }}
                            />
                            {ATLessons.length > 1 ? (
                              <Icon
                                aria-label="Delete"
                                iconName="Delete"
                                data-index={les.Index}
                                className={ATIconStyleClass.delete}
                                onClick={(e) => {
                                  let dataindex =
                                    e.target["getAttribute"]("data-index");
                                  let tempArr = [...ATLessons];
                                  let targetArrIndex = tempArr.findIndex(
                                    (dt) => dt.Index == dataindex
                                  );
                                  tempArr.splice(targetArrIndex, 1);
                                  setATLessons([...tempArr]);
                                  ATOnchange(
                                    "Lessons",
                                    tempArr.map((user) => user.Text).join(";")
                                  );
                                }}
                              />
                            ) : null}
                          </div>
                        </div>
                      </div>
                    );
                  })
                : null}
            </div>

            <div className={styles.apModalBoxButtonSection}>
              <button
                className={styles.apModalBoxSubmitBtn}
                onClick={(_) => {
                  ATValidationFunction();
                }}
                style={{ display: "flex" }}
              >
                {ATModalBoxPopup.type == "New" ? (
                  ATOnSubmitLoader ? (
                    <Spinner />
                  ) : (
                    <span>
                      <Icon
                        iconName="Save"
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      {"Submit"}
                    </span>
                  )
                ) : ATOnSubmitLoader ? (
                  <Spinner />
                ) : (
                  <span>
                    <Icon
                      iconName="Save"
                      style={{ position: "relative", top: 3, left: -8 }}
                    />
                    {"Update"}
                  </span>
                )}
              </button>
              <button
                className={styles.apModalBoxBackBtn}
                onClick={(_) => {
                  // Non dropdown value remove
                  let isOriginalData_Project = ProjectOrProductDetails.filter(
                    (arr) => {
                      return (arr.Type =
                        "Project" &&
                        arr.Key == ATModalBoxPopup["value"]["Project"]);
                    }
                  );
                  let isOriginalData_Product = ProjectOrProductDetails.filter(
                    (arr) => {
                      return (arr.Type =
                        "Product" &&
                        arr.Key == ATModalBoxPopup["value"]["Product"]);
                    }
                  );
                  if (
                    isOriginalData_Project.length == 0 &&
                    ATModalBoxPopup["value"]["Project"]
                  ) {
                    ATModalBoxDropDownOptions.Project.pop();
                  }
                  if (
                    isOriginalData_Product.length == 0 &&
                    ATModalBoxPopup["value"]["Product"]
                  ) {
                    ATModalBoxDropDownOptions.Product.pop();
                  }

                  setATModalBoxDropDownOptions({
                    ...ATModalBoxDropDownOptions,
                  });

                  setATModalBoxPopup({
                    type: ATModalBoxPopup.type,
                    visible: false,
                    value: {},
                  });
                }}
              >
                <span>
                  {" "}
                  <Icon
                    iconName="Cancel"
                    style={{ position: "relative", top: 3, left: -8 }}
                  />
                  Close
                </span>
              </button>
            </div>
          </div>
        </Modal>
      </div>
    </div>
  );
};

export default ActivityTemplate;
