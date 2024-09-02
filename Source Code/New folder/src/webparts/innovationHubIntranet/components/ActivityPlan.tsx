import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  NormalPeoplePicker,
  Modal,
  IModalStyles,
  DatePicker,
  IDatePickerStyles,
  Spinner,
  TooltipHost,
  TooltipOverflowMode,
  IColumn,
  TextField,
  ITextFieldStyles,
} from "@fluentui/react";

import Service from "../components/Services";

import MUITextField from "@material-ui/core/TextField";
import Autocomplete from "@material-ui/lab/Autocomplete";
import "../ExternalRef/styleSheets/Styles.css";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

let columnSortArr = [];
let columnSortMasterArr = [];
let atpedittemplate;
let ProjectOrProductDetails = [];
let DateListFormat = "DD/MM/YYYY";

const ActivityPlan = (props: any) => {
  // Variable-Declaration-Section Starts
  const sharepointWeb = Web(props.URL);
  const ListName = "Activity Plan";
  const TemplateListName = "Activity Template";

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  let Ap_AnnualPlanId = props.AnnualPlanId;
  const ListNameURL = props.WeblistURL;

  const atpAllitems = [];
  const allPeoples = props.peopleList;
  const atpColumns = [
    {
      key: "Types",
      name: "Types",
      fieldName: "Types",
      minWidth: 100,
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
      key: "Area",
      name: "Area/Stream",
      fieldName: "Area",
      minWidth: 75,
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
      key: "Product",
      name: "Product(Program)",
      fieldName: "Product",
      minWidth: 100,
      maxWidth: 200,
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
      key: "Project",
      name: "Project",
      fieldName: "Project",
      minWidth: 100,
      maxWidth: 200,
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
      key: "ProductCode",
      name: "Code",
      fieldName: "ProductCode",
      minWidth: 50,
      maxWidth: 70,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.ProductCode}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.ProductCode}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "ActivityPlanName",
      name: "Activity plan name",
      fieldName: "ActivityPlanName",
      minWidth: 150,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.ActivityPlanName}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.ActivityPlanName}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "TemplateName",
      name: "Template",
      fieldName: "TemplateName",
      minWidth: 100,
      maxWidth: 200,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.TemplateName}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.TemplateName}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Status",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 120,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) => (
        <>
          {item.Status == "Completed" ? (
            <div className={atpStatusStyleClass.completed}>{item.Status}</div>
          ) : item.Status == "Scheduled" ? (
            <div className={atpStatusStyleClass.scheduled}>{item.Status}</div>
          ) : item.Status == "On schedule" ? (
            <div className={atpStatusStyleClass.onSchedule}>{item.Status}</div>
          ) : item.Status == "Behind schedule" ? (
            <div className={atpStatusStyleClass.behindScheduled}>
              {item.Status}
            </div>
          ) : item.Status == "On hold" ? (
            <div className={atpStatusStyleClass.Onhold}>{item.Status}</div>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "CompletedDate",
      name: "Completed on",
      fieldName: "CompletedDate",
      minWidth: 100,
      maxWidth: 110,
      onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
        _onColumnClick(ev, column);
      },
      onRender: (item) =>
        item.CompletedDate != null ? item.CompletedDate : "",
    },
    {
      key: "ADP",
      name: "AP",
      fieldName: "ADP",
      minWidth: 30,
      maxWidth: 30,

      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={atpIconStyleClass.link}
            onClick={() => {
              atpEdit.length > 0
                ? ""
                : props.handleclick("ActivityDeliveryPlan", item.ID, "ATP");
            }}
          />
        </>
      ),
    },
    {
      key: "APB",
      name: "PB",
      fieldName: "APB",
      minWidth: 30,
      maxWidth: 30,

      onRender: (item) => (
        <>
          <Icon
            iconName="Link12"
            className={atpIconStyleClass.link}
            onClick={() => {
              atpEdit.length > 0
                ? ""
                : props.handleclick("ActivityProductionBoard", item.ID, "ATP");
            }}
          />
        </>
      ),
    },
    {
      key: "Action",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 100,

      onRender: (item) => (
        <>
          <Icon
            iconName="Edit"
            className={atpIconStyleClass.edit}
            onClick={() => {
              let selectedTemplate = item.TemplateName.slice(
                0,
                item.TemplateName.split("~")[0].length - 1
              );
              // let templateDetails = { ...atpTemplatesDetails };
              let templateDetails = atpedittemplate;
              templateDetails.templateOptns = [
                {
                  key: selectedTemplate,
                  text: selectedTemplate,
                },
              ];

              if (
                templateDetails.projectOptns.findIndex((_projectDrpDwn) => {
                  return _projectDrpDwn.key == item.Project;
                }) == -1 &&
                item.Project
              ) {
                templateDetails.projectOptns.push({
                  key: item.Project,
                  text: item.Project,
                });
              }

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
                atpDropDownOptions.productAllOptns.push({
                  key: item.Product,
                  text: item.Product,
                });
              }
              if (isOriginalData_Project.length == 0) {
                atpDropDownOptions.projectAllOptns.push({
                  key: item.Project,
                  text: item.Project,
                });
              }
              setAtpDropDownOptions({ ...atpDropDownOptions });

              setAtpLoader("noLoader");
              setAtpTemplatesDetails({ ...templateDetails });
              atpActivityResponseArrGenerator("data", [item]);
              setAtpEdit([item]);
              setAtpPlannerTemplate(selectedTemplate);
              setAtpPlannerProduct(item.Product);
              setAtpPlannerProject(item.Project);
              setatpCopy({
                IsCopy: false,
                Project: item.Project,
                Product: item.Product,
                ActivityPlanName: item.ActivityPlanName,
                IsValidation: false,
              });
            }}
          />
          <Icon
            iconName="Copy"
            className={atpIconStyleClass.copy}
            onClick={() => {
              let selectedTemplate = item.TemplateName.slice(
                0,
                item.TemplateName.split("~")[0].length - 1
              );
              // let templateDetails = { ...atpTemplatesDetails };
              let templateDetails = atpedittemplate;
              templateDetails.templateOptns = [
                {
                  key: selectedTemplate,
                  text: selectedTemplate,
                },
              ];
              if (
                templateDetails.projectOptns.findIndex((_projectDrpDwn) => {
                  return _projectDrpDwn.key == item.Project;
                }) == -1 &&
                item.Project
              ) {
                templateDetails.projectOptns.push({
                  key: item.Project,
                  text: item.Project,
                });
              }

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
                atpDropDownOptions.productAllOptns.push({
                  key: item.Product,
                  text: item.Product,
                });
              }
              if (isOriginalData_Project.length == 0) {
                atpDropDownOptions.projectAllOptns.push({
                  key: item.Project,
                  text: item.Project,
                });
              }
              setAtpDropDownOptions({ ...atpDropDownOptions });

              setAtpLoader("noLoader");
              setAtpTemplatesDetails({ ...templateDetails });
              atpActivityResponseArrGenerator("data", [item]);
              setAtpEdit([item]);
              setatpCopy({
                IsCopy: true,
                Project: item.Project,
                Product: item.Product,
                ActivityPlanName: item.ActivityPlanName,
                IsValidation: false,
              });
              setAtpPlannerTemplate(selectedTemplate);
              setAtpPlannerProduct(item.Product);
              setAtpPlannerProject(item.Project);
            }}
          />
          <Icon
            iconName="Delete"
            className={atpIconStyleClass.delete}
            onClick={() => {
              atpEdit.length > 0
                ? ""
                : setAtpDeletePopup({ targetID: item.ID, condition: true });
            }}
          />
        </>
      ),
    },
  ];
  const atpModalBoxColumns = [
    {
      key: "Lesson",
      name: "Section",
      fieldName: "Lesson",
      minWidth: 180,
      maxWidth: 180,
    },
    {
      key: "StartDate",
      name: "Start date",
      fieldName: "StartDate",
      minWidth: 150,
      maxWidth: 150,

      onRender: (item, index: number) => (
        <>
          <DatePicker
            placeholder="Select a start date"
            formatDate={dateFormater}
            styles={
              atpActivityResponseData[index][`dateError${index}`] == true
                ? {
                    textField: {
                      selectors: {
                        ".ms-TextField-fieldGroup": {
                          border: "none !important",
                        },
                      },
                    },
                    readOnlyTextField: {
                      color: "#d0342c !important",
                      border: "2px solid #d0342c !important",
                    },
                    icon: {
                      fontSize: 18,
                      color: "#d0342c",
                    },
                  }
                : {
                    readOnlyTextField: {
                      color: "#7C7C7C !important",
                    },
                    icon: {
                      fontSize: 18,
                      color: "#7C7C7C",
                    },
                  }
            }
            value={
              atpActivityResponseData[index][`startDate${index}`]
                ? atpActivityResponseData[index][`startDate${index}`]
                : new Date()
            }
            onSelectDate={(value: any) => {
              atpActivityResponseHandler(index, `startDate${index}`, value);
            }}
          />
        </>
      ),
    },
    {
      key: "EndDate",
      name: "End date",
      fieldName: "EndDate",
      minWidth: 150,
      maxWidth: 150,

      onRender: (item, index: number) => (
        <>
          <DatePicker
            placeholder="Select a end date"
            formatDate={dateFormater}
            styles={
              atpActivityResponseData[index][`dateError${index}`] == true
                ? {
                    textField: {
                      selectors: {
                        ".ms-TextField-fieldGroup": {
                          border: "none !important",
                        },
                      },
                    },
                    readOnlyTextField: {
                      color: "#d0342c !important",
                      border: "2px solid #d0342c !important",
                    },
                    icon: {
                      fontSize: 18,
                      color: "#d0342c",
                    },
                  }
                : {
                    readOnlyTextField: {
                      color: "#7C7C7C !important",
                    },
                    icon: {
                      fontSize: 18,
                      color: "#7C7C7C",
                    },
                  }
            }
            value={
              atpActivityResponseData[index][`endDate${index}`]
                ? atpActivityResponseData[index][`endDate${index}`]
                : new Date()
            }
            onSelectDate={(value: any) => {
              atpActivityResponseHandler(index, `endDate${index}`, value);
            }}
          />
        </>
      ),
    },
    {
      key: "Developer",
      name: "Developer",
      fieldName: "Developer",
      minWidth: 220,
      maxWidth: 220,

      onRender: (item, index: number) => (
        <>
          <NormalPeoplePicker
            onResolveSuggestions={GetUserDetails}
            itemLimit={1}
            selectedItems={allPeoples.filter((people) => {
              return (
                people.ID ==
                (atpActivityResponseData[index][`developer${index}`]
                  ? atpActivityResponseData[index][`developer${index}`]
                  : null)
              );
            })}
            onChange={(selectedUser) => {
              atpActivityResponseHandler(
                index,
                `developer${index}`,
                selectedUser[0] ? selectedUser[0]["ID"] : null
              );
            }}
          />
        </>
      ),
    },
  ];
  const atpStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
  });
  const atpStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      atpStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      atpStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      atpStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      atpStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      atpStatusStyle,
    ],
  });
  const atpDrpDwnOptns = {
    typesOptns: [{ key: "All", text: "All" }],
    areaOptns: [{ key: "All", text: "All" }],
    productOptns: [{ key: "All", text: "All" }],
    projectOptns: [{ key: "All", text: "All" }],
    productAllOptns: [],
    projectAllOptns: [],
  };
  const atpFilterKeys = {
    type: "All",
    area: "All",
    product: "All",
    project: "All",
    code: "",
    template: "",
  };
  const ATTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: 300 },
    field: { backgroundColor: "white", fontSize: 12 },
  };
  let currentpage = 1;
  let totalPageItems = 10;
  // Variable-Declaration-Section Ends
  // Styles-Section Starts

  const atpDetailsListStyles: Partial<IDetailsListStyles> = {
    // root: {},
    // headerWrapper: {},
    // contentWrapper: {
    //   ".ms-DetailsRow-cell": {
    //     paddingBottom: "0 !important",
    //   },
    // },
    root: {
      selectors: {
        ".ms-DetailsRow-fields": { minHeight: 38 },
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          alignItems: "start",
          ".ms-DetailsRow-cell": {
            height: 38,
            minHeight: 38,
            padding: "11px 12px",
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
  const atpModalBoxDetailsListStyles: Partial<IDetailsListStyles> = {
    root: {
      width: 960,
      overflowX: "none",
      selectors: {
        ".ms-DetailsRow-cell": {
          height: 45,
        },
      },
    },
    headerWrapper: {},
    contentWrapper: { height: 140, overflowX: "hidden", overflowY: "auto" },
  };
  const atpLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      fontSize: 13,
      color: "#323130",
    },
  };
  const atpModalBoxLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 165,
      marginTop: 5,
      marginRight: 10,
      marginLeft: 15,
      fontSize: 13,
      color: "#323130",
    },
  };
  const atpDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
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
  const atpActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 165,
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
  const atpSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
      marginTop: "3px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const atpActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 165,
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
  const atpDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: 165,
      marginRight: 15,
      marginTop: 5,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    wrapper: {
      borderRadius: "4px",
      ".ms-TextField-fieldGroup": {
        border: "none",
      },
      ".ms-TextField-field": {
        borderRadius: "4px !important",
      },
      ".readOnlyPlaceholder-203": {
        color: "#7C7C7C !important",
      },
    },
    readOnlyTextField: {
      backgroundColor: "#F5F5F7 !important",
      fontSize: 12,
      border: "1px solid #E8E8EA !important",
      borderRadius: 4,
    },
    icon: {
      fontSize: 18,
      color: "#7C7C7C",
    },
  };
  const atpActiveDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: 165,
      marginRight: 15,
      marginTop: 5,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    wrapper: {
      borderRadius: "4px",
      ".ms-TextField-fieldGroup": {
        border: "none",
      },
      ".ms-TextField-field": {
        borderRadius: "4px !important",
      },
      ".readOnlyPlaceholder-203": {
        color: "#038387 !important",
      },
    },
    readOnlyTextField: {
      backgroundColor: "#F5F5F7 !important",
      fontSize: 12,
      border: "2px solid #038387 !important",
      borderRadius: 4,
      color: "#038387",
      fontWeight: 600,
    },
    icon: {
      fontSize: 18,
      color: "#038387",
      fontWeight: 600,
    },
  };
  const atpModalStyles: Partial<IModalStyles> = {
    root: { borderRadius: "none" },
    main: {
      width: 1000,
      minHeight: 550,
      margin: 10,
      padding: "20px 10px",
      display: "flex",
      flexDirection: "column",
    },
  };
  const atpCopyModalBoxDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "200px",
      maxWidth: "300px",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const atpModalBoxDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      marginTop: 5,
      marginRight: 10,
      marginLeft: 15,
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
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "200px",
      maxWidth: "300px",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const atpModalBoxActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      marginTop: 5,
      marginRight: 10,
      marginLeft: 15,
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
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "200px",
      maxWidth: "300px",
    },
    caretDown: { fontSize: 14, color: "#038387", fontWeight: 600 },
  };
  const atpModalBoxReadOnlyDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      marginTop: 5,
      marginRight: 10,
      marginLeft: 15,
      backgroundColor: "#F5F5F7",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "2px solid #7C7C7C",
      borderRadius: 4,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "200px",
      maxWidth: "300px",
    },
    caretDown: { fontSize: 14, color: "#7C7C7C", display: "none" },
  };
  const atpIconStyleClass = mergeStyleSets({
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
    linkDisabled: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#ababab",
        cursor: "not-allowed",
      },
    ],
    edit: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    save: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#36b04b",
        cursor: "pointer",
      },
    ],
    saveDisabled: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        color: "#ababab",
        cursor: "not-allowed",
      },
    ],
    delete: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#CB1E06",
        cursor: "pointer",
      },
    ],
    copy: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#2392B2",
        cursor: "pointer",
      },
    ],
    deleteDisabled: [
      {
        fontSize: 18,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#ababab",
        cursor: "not-allowed",
      },
    ],
    close: [
      {
        fontSize: 15,
        height: 14,
        width: 17,
        marginLeft: 10,
        color: "#CB1E06",
        cursor: "pointer",
      },
    ],
    linkPB: [
      {
        fontSize: 18,
        height: 16,
        width: 19,
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
        marginTop: 36.5,
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
        marginTop: "-7px",
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
  const generalStyles = mergeStyleSets({
    titleLabel: {
      color: "#2392B2 !important",
      fontWeight: "500",
      fontSize: 17,
    },
    inputLabel: {
      color: "#2392B2 !important",
      display: "block",
      fontWeight: "500",
      margin: "5px 0",
    },
    inputValue: {
      color: "#000",
      fontWeight: "500",
      fontSize: 13,
    },
    inputField: {
      margin: "10px 0",
    },
    dateValidationErrorLabel: {
      color: "#d0342c !important",
      fontWeight: 600,
      marginTop: 40,
      marginLeft: 20,
    },
    dateGridValidationErrorLabel: {
      color: "#d0342c !important",
      fontWeight: 600,
      marginLeft: 20,
    },
  });
  // Styles-Section Ends
  // States-Declaration Starts
  const [atpReRender, setAtpReRender] = useState(true);
  // const [atpPeopleList, setAtpPeopleList] = useState(allPeoples);
  const [currentUser, setCurrentUser] = useState({});
  const [atpUnsortMasterData, setAtpUnsortMasterData] = useState(atpAllitems);
  const [atpMasterData, setAtpMasterData] = useState(atpAllitems);
  const [atpData, setAtpData] = useState(atpAllitems);
  const [atpDisplayData, setAtpDisplayData] = useState([]);
  const [atpActivityResponseData, setAtpActivityResponseData] = useState([]);
  const [atpcurrentPage, setAtpCurrentPage] = useState(currentpage);
  const [atpDropDownOptions, setAtpDropDownOptions] = useState(atpDrpDwnOptns);
  const [atpFilters, setAtpFilters] = useState(atpFilterKeys);
  const [atpEdit, setAtpEdit] = useState([]);
  const [atpCopy, setatpCopy] = useState({
    IsCopy: false,
    Project: "",
    Product: "",
    ActivityPlanName: "",
    IsValidation: false,
  });
  const [atpLessonData, setAtpLessonData] = useState([]);
  const [atpAddPlannerPopup, setAtpAddPlannerPopup] = useState(false);
  const [atpDeletePopup, setAtpDeletePopup] = useState({
    targetID: null,
    condition: false,
  });
  const [atpTemplatesDetails, setAtpTemplatesDetails] = useState({
    data: [],
    templateOptns: [],
    projectOptns: [],
    productOptns: [],
  });
  const [atpPlannerTemplate, setAtpPlannerTemplate] = useState(null);
  const [atpPlannerProject, setAtpPlannerProject] = useState(null);
  const [atpPlannerProduct, setAtpPlannerProduct] = useState(null);
  const [atpLoader, setAtpLoader] = useState("noLoader");
  const [atpMasterColumns, setAtpMasterColumns] = useState(atpColumns);

  const options = [
    "All",
    "Option 2",
    "Option 3",
    "Option 4",
    "Option 5",
    "Option 6",
    "Option 7",
    "Option 8",
    "Option 9",
    "Option 10",
  ];
  const [value, setValue] = useState("");
  const [inputValue, setInputValue] = useState("");
  // States-Declaration Ends
  //Function-Section Starts

  const generateExcel = () => {
    let arrExport = atpData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Types", key: "Types", width: 25 },
      { header: "Area/Stream", key: "Area", width: 25 },
      { header: "Product", key: "Product", width: 25 },
      { header: "Project", key: "Project", width: 25 },
      { header: "Code", key: "Code", width: 25 },
      { header: "Activity plan name", key: "ActivityPlanName", width: 30 },
      { header: "Template", key: "Template", width: 60 },
      { header: "Status", key: "Status", width: 20 },
      { header: "Completed on", key: "CompletedDate", width: 40 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        Types: item.Types ? item.Types : "",
        Area: item.Area ? item.Area : "",
        Product: item.Product ? item.Product : "",
        Project: item.Project ? item.Project : "",
        Code: item.Code ? item.Code : "",
        Template: item.TemplateName ? item.TemplateName : "",
        Status: item.Status ? item.Status : "",
        CompletedDate: item.CompletedDate ? item.CompletedDate : "",
        ActivityPlanName: item.ActivityPlanName ? item.ActivityPlanName : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1", "I1"].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `ActivityPlan-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };

  const atpGetCurrentUserDetails = () => {
    sharepointWeb.currentUser
      .get()
      .then((user) => {
        let atpCurrentUser = {
          Name: user.Title,
          Email: user.Email,
          Id: user.Id,
        };
        setCurrentUser(atpCurrentUser);
      })
      .catch((err) => {
        atpErrorFunction(err, "atpGetCurrentUserDetails");
      });
  };
  const getCurrentAPData = () => {
    let _apCurrentData = [];
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/PlannedEndDate"
      )
      .expand(
        "ProjectOwner",
        "ProjectLead",
        "Master_x0020_Project",
        "FieldValuesAsText"
      )
      .filter("ID eq '" + Ap_AnnualPlanId + "' ")
      .top(5000)
      .get()
      .then(async (items) => {
        items.forEach((item) => {
          _apCurrentData.push({
            ID: item.ID,
            Title: item.Title
              ? item.Title +
                " " +
                (item.ProjectVersion ? item.ProjectVersion : "V1")
              : "",
            TypeofProject: item.ProjectType,
            ProductId: item.Master_x0020_ProjectId,
            ProductName: item.Master_x0020_Project
              ? item.Master_x0020_Project.Title +
                " " +
                (item.ProductVersion ? item.ProductVersion : "V1")
              : "",
            DeveloperId: item.ProjectLeadId ? item.ProjectLeadId[0] : null,
            ProjectOwnerId: item.ProjectOwnerId,
            DeveloperEmail: item.ProjectLead ? item.ProjectLead[0].EMail : null,
            ProjectOwnerEmail: item.ProjectOwner
              ? item.ProjectOwner.EMail
              : null,
            DeveloperName: item.ProjectLead ? item.ProjectLead[0].Title : null,
            ProjectOwnerName: item.ProjectOwner
              ? item.ProjectOwner.Title
              : null,
            StartDate: moment(
              item["FieldValuesAsText"].StartDate,
              DateListFormat
            ).format(DateListFormat),
            PlannedEndDate: moment(
              item["FieldValuesAsText"].PlannedEndDate,
              DateListFormat
            ).format(DateListFormat),
            AllocatedHours: item.AllocatedHours ? item.AllocatedHours : 0,
            BA: item.BusinessArea,
            MultipleDeveloper: item.ProjectLead ? item.ProjectLead : null,
          });
        });
        console.log(_apCurrentData[0]);
        atpGetLinkedData(_apCurrentData[0]);
      })
      .catch((err) => {
        atpErrorFunction(err, "getCurrentAPData");
      });
  };
  const atpGetLinkedData = (apData) => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.select("*", "FieldValuesAsText/CompletedDate")
      .expand("FieldValuesAsText")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item) => {
          atpAllitems.push({
            ID: item.Id ? item.Id : "",
            TemplateName: item.Title ? item.Title : "",
            Lessons: item.Lessons ? item.Lessons : "",
            Project: item.Project
              ? item.Project +
                " " +
                (item.ProjectVersion ? item.ProjectVersion : "V1")
              : "",
            ActivityPlanName: item.ActivityPlanName
              ? item.ActivityPlanName
              : "",
            Types: item.Types ? item.Types : "",
            Area: item.Area ? item.Area : "",
            Product: item.Product
              ? item.Product +
                " " +
                (item.ProductVersion ? item.ProductVersion : "V1")
              : "",
            ProductCode: item.ProductCode ? item.ProductCode : "",
            Status: item.Status ? item.Status : "",
            CompletedDate: item.CompletedDate
              ? moment(
                  item["FieldValuesAsText"].CompletedDate,
                  DateListFormat
                ).format(DateListFormat)
              : null,
          });
        });
        console.log(atpAllitems);

        setAtpUnsortMasterData([...atpAllitems]);
        columnSortMasterArr = atpAllitems;
        setAtpMasterData([...atpAllitems]);
        atpGetAllOptions(atpAllitems, apData);

        // setAtpFilters({ ...FilterKeys });
        // columnSortArr = atpAllitems;
        // setAtpData([...atpAllitems]);
        // paginateFunction(1, atpAllitems);
        // setAtpLoader("noLoader");
      })
      .catch((err) => {
        atpErrorFunction(err, "atpGetLinkedData");
      });
  };
  const atpGetData = () => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.select("*", "FieldValuesAsText/CompletedDate")
      .expand("FieldValuesAsText")
      .orderBy("Modified", false)
      .top(5000)
      .get()
      .then((items) => {
        items.forEach((item) => {
          atpAllitems.push({
            ID: item.Id ? item.Id : "",
            TemplateName: item.Title ? item.Title : "",
            Lessons: item.Lessons ? item.Lessons : "",
            Project: item.Project
              ? item.Project +
                " " +
                (item.ProjectVersion ? item.ProjectVersion : "V1")
              : "",
            Types: item.Types ? item.Types : "",
            ActivityPlanName: item.ActivityPlanName
              ? item.ActivityPlanName
              : "",
            Area: item.Area ? item.Area : "",
            Product: item.Product
              ? item.Product +
                " " +
                (item.ProductVersion ? item.ProductVersion : "V1")
              : "",
            ProductCode: item.ProductCode ? item.ProductCode : "",
            Status: item.Status ? item.Status : "",
            CompletedDate: item.CompletedDate
              ? moment(
                  item["FieldValuesAsText"].CompletedDate,
                  DateListFormat
                ).format(DateListFormat)
              : null,
          });
        });
        console.log(atpAllitems);
        atpGetAllOptions(atpAllitems, null);
        paginateFunction(1, atpAllitems);

        setAtpUnsortMasterData([...atpAllitems]);
        columnSortArr = atpAllitems;
        setAtpData([...atpAllitems]);
        columnSortMasterArr = atpAllitems;
        setAtpMasterData([...atpAllitems]);
        setAtpLoader("noLoader");
      })
      .catch((err) => {
        atpErrorFunction(err, "atpGetData");
      });
  };
  const atpGetTemplates = () => {
    sharepointWeb.lists
      .getByTitle(TemplateListName)
      .items.top(5000)
      .get()
      .then((items) => {
        let _templatesData = [];
        let _templatesDrpDwns = [];
        let _projectDrpDwns = [];
        let _productDrpDwns = [];

        items.forEach((item) => {
          let lessons_str_to_arr = item.Lessons
            ? item.Lessons.split(";")
            : null;
          let tempLessonArr = [];
          if (lessons_str_to_arr) {
            lessons_str_to_arr.forEach((lesson) => {
              tempLessonArr.push({ Lesson: lesson });
            });
          }
          let PrdValue = item.Product
            ? item.Product +
              " " +
              (item.ProductVersion ? item.ProductVersion : "V1")
            : "";
          let PrjValue = item.Project
            ? item.Project +
              " " +
              (item.ProjectVersion ? item.ProjectVersion : "V1")
            : "";

          _templatesData.push({
            TemplateName: item.Title ? item.Title : "",
            Project: PrjValue,
            Types: item.Types ? item.Types : "",
            ActivityPlanName: item.ActivityPlanName
              ? item.ActivityPlanName
              : "",
            Area: item.Area ? item.Area : "",
            Product: PrdValue,
            ProductCode: item.Code ? item.Code : "",
            LessonList: tempLessonArr,
            Status: "Scheduled",
            CompletedDate: null,
            IsDeleted: item.IsDeleted,
          });

          if (
            _productDrpDwns.findIndex((_productDrpDwn) => {
              return _productDrpDwn.key == PrdValue;
            }) == -1 &&
            PrdValue &&
            item.IsDeleted != true
          ) {
            _productDrpDwns.push({ key: PrdValue, text: PrdValue });
          }

          if (
            _projectDrpDwns.findIndex((_projectDrpDwn) => {
              return _projectDrpDwn.key == PrjValue;
            }) == -1 &&
            PrjValue &&
            item.IsDeleted != true
          ) {
            _projectDrpDwns.push({ key: PrjValue, text: PrjValue });
          }
        });

        setAtpTemplatesDetails({
          data: _templatesData,
          projectOptns: _projectDrpDwns,
          templateOptns: _templatesDrpDwns,
          productOptns: _productDrpDwns,
        });
        atpedittemplate = {
          data: _templatesData,
          templateOptns: _templatesDrpDwns,
          productOptns: _productDrpDwns,
          projectOptns: _projectDrpDwns,
        };
      })
      .catch((err) => {
        atpErrorFunction(err, "atpGetTemplates");
      });
  };
  const atpGetAllOptions = (allItems: any, apData: any) => {
    ProjectOrProductDetails = [];
    //Product Choices

    const _sortFilterKeys = (a, b) => {
      if (a.text.toLowerCase() < b.text.toLowerCase()) {
        return -1;
      }
      if (a.text.toLowerCase() > b.text.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.top(5000)
      .orderBy("Modified", false)
      .get()
      .then((Items) => {
        Items.forEach((arr) => {
          if (
            atpDrpDwnOptns.projectAllOptns.findIndex((prj) => {
              return prj.key == arr.Title;
            }) == -1 &&
            arr.Title
          ) {
            atpDrpDwnOptns.projectAllOptns.push({
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
        atpDrpDwnOptns.projectAllOptns.sort(_sortFilterKeys);

        sharepointWeb.lists
          .getByTitle("Master Product List")
          .items.filter("IsDeleted ne 1")
          .top(5000)
          .get()
          .then((allProducts) => {
            allProducts.forEach((product) => {
              if (product.Title != null) {
                if (
                  atpDrpDwnOptns.productAllOptns.findIndex((productOptn) => {
                    return productOptn.key == product.Title;
                  }) == -1
                ) {
                  if (product.Title != "Not Sure") {
                    atpDrpDwnOptns.productAllOptns.push({
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
            atpDrpDwnOptns.productAllOptns.sort(_sortFilterKeys);
            atpDrpDwnOptns.productAllOptns.unshift({
              key: "Not Sure V1",
              text: "Not Sure V1",
            });

            allItems.forEach((item: any) => {
              if (
                atpDrpDwnOptns.typesOptns.findIndex((typesOptn) => {
                  return typesOptn.key == item.Types;
                }) == -1 &&
                item.Types
              ) {
                atpDrpDwnOptns.typesOptns.push({
                  key: item.Types,
                  text: item.Types,
                });
              }

              if (
                atpDrpDwnOptns.areaOptns.findIndex((areaOptn) => {
                  return areaOptn.key == item.Area;
                }) == -1 &&
                item.Area
              ) {
                atpDrpDwnOptns.areaOptns.push({
                  key: item.Area,
                  text: item.Area,
                });
              }

              if (
                atpDrpDwnOptns.productOptns.findIndex((productOptn) => {
                  return productOptn.key == item.Product;
                }) == -1 &&
                item.Product
              ) {
                atpDrpDwnOptns.productOptns.push({
                  key: item.Product,
                  text: item.Product,
                });
              }

              if (
                atpDrpDwnOptns.projectOptns.findIndex((projectOptn) => {
                  return projectOptn.key == item.Project;
                }) == -1 &&
                item.Project
              ) {
                atpDrpDwnOptns.projectOptns.push({
                  key: item.Project,
                  text: item.Project,
                });
              }
            });

            let unsortedFilterKeys = atpSortingFilterKeys(atpDrpDwnOptns);
            setAtpDropDownOptions({ ...unsortedFilterKeys });

            if (Ap_AnnualPlanId && apData) {
              let FilterKeys = {
                type: "All",
                area: "All",
                project: apData.Title,
                product: apData.ProductName,
                code: "",
                template: "",
              };

              atpOnloadListFilter(atpAllitems, FilterKeys);
            }
          })
          .catch((err) => {
            atpErrorFunction(err, "atpGetAllOptions-master product list");
          });
      })
      .catch((err) => {
        atpErrorFunction(err, "atpGetAllOptions-Project");
      });
  };
  const atpSortingFilterKeys = (unsortedFilterKeys: any) => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    unsortedFilterKeys.typesOptns.shift();
    unsortedFilterKeys.typesOptns.sort(sortFilterKeys);
    unsortedFilterKeys.typesOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.areaOptns.shift();
    unsortedFilterKeys.areaOptns.sort(sortFilterKeys);
    unsortedFilterKeys.areaOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.productOptns.shift();
    unsortedFilterKeys.productOptns.sort(sortFilterKeys);
    unsortedFilterKeys.productOptns.unshift({ key: "All", text: "All" });

    unsortedFilterKeys.projectOptns.shift();
    unsortedFilterKeys.projectOptns.sort(sortFilterKeys);
    unsortedFilterKeys.projectOptns.unshift({ key: "All", text: "All" });

    return unsortedFilterKeys;
  };
  const atpListFilter = (key: string, option: any) => {
    let arrBeforeFilter = [...atpMasterData];
    let tempFilterKeys = { ...atpFilters };
    tempFilterKeys[key] = option;

    if (tempFilterKeys.type != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Types == tempFilterKeys.type;
      });
    }

    if (tempFilterKeys.area != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Area == tempFilterKeys.area;
      });
    }
    if (tempFilterKeys.product != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Product == tempFilterKeys.product;
      });
    }

    if (tempFilterKeys.project != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Project == tempFilterKeys.project;
      });
    }

    if (tempFilterKeys.template) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.TemplateName.toLowerCase().includes(
          tempFilterKeys.template.toLowerCase()
        );
      });
    }

    if (tempFilterKeys.code) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.ProductCode.toLowerCase().includes(
          tempFilterKeys.code.toLowerCase()
        );
      });
    }

    paginateFunction(1, arrBeforeFilter);

    columnSortArr = arrBeforeFilter;
    setAtpData([...columnSortArr]);
    setAtpFilters({ ...tempFilterKeys });
  };
  const atpOnloadListFilter = (data, filterskeys) => {
    let arrBeforeFilter = [...data];
    let tempFilterKeys = { ...filterskeys };

    // if (tempFilterKeys.type != "All") {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.Types == tempFilterKeys.type;
    //   });
    // }

    // if (tempFilterKeys.area != "All") {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.Area == tempFilterKeys.area;
    //   });
    // }
    if (tempFilterKeys.product != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Product == tempFilterKeys.product;
      });
    }

    if (tempFilterKeys.project != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Project == tempFilterKeys.project;
      });
    }

    // if (tempFilterKeys.template) {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.TemplateName.toLowerCase().includes(
    //       tempFilterKeys.template.toLowerCase()
    //     );
    //   });
    // }

    // if (tempFilterKeys.code) {
    //   arrBeforeFilter = arrBeforeFilter.filter((arr) => {
    //     return arr.ProductCode.toLowerCase().includes(
    //       tempFilterKeys.code.toLowerCase()
    //     );
    //   });
    // }

    paginateFunction(1, arrBeforeFilter);

    columnSortArr = arrBeforeFilter;
    setAtpData([...columnSortArr]);

    columnSortArr.length > 0
      ? setAtpFilters({ ...tempFilterKeys })
      : setAtpFilters({ ...atpFilterKeys });

    setAtpLoader("noLoader");
  };
  const atpListFilterbyData = (_masterData: any) => {
    let arrBeforeFilter = [..._masterData];
    let tempFilterKeys = { ...atpFilters };

    if (tempFilterKeys.type != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Types == tempFilterKeys.type;
      });
    }

    if (tempFilterKeys.area != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Area == tempFilterKeys.area;
      });
    }
    if (tempFilterKeys.product != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Product == tempFilterKeys.product;
      });
    }

    if (tempFilterKeys.project != "All") {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Project == tempFilterKeys.project;
      });
    }

    if (tempFilterKeys.template) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Lessons.toLowerCase().includes(
          tempFilterKeys.template.toLowerCase()
        );
      });
    }

    if (tempFilterKeys.code) {
      arrBeforeFilter = arrBeforeFilter.filter((arr) => {
        return arr.Lessons.toLowerCase().includes(
          tempFilterKeys.code.toLowerCase()
        );
      });
    }

    return arrBeforeFilter;
  };
  const findRefID = (option: String) => {
    const sortFilterKeys = (b, a) => {
      if (a.TemplateName.toLowerCase() < b.TemplateName.toLowerCase()) {
        return -1;
      }
      if (a.TemplateName.toLowerCase() > b.TemplateName.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    let filteredArr = atpMasterData.filter((data) => {
      return data.TemplateName.includes(option);
    });

    if (filteredArr.length > 0) {
      if (filteredArr.length == 1) {
        return parseInt(filteredArr[0].TemplateName.split("~")[1]) + 1;
      } else {
        return (
          parseInt(
            filteredArr.sort(sortFilterKeys)[0].TemplateName.split("~")[1]
          ) + 1
        );
      }
    } else {
      return 1;
    }
  };
  const atpActivityResponseArrGenerator = (key: string, data: any) => {
    if (key == "template") {
      let tempActivityResponseData = [];
      let newLessonArr = atpTemplatesDetails.data.filter((temp) => {
        return temp.TemplateName == data;
      })[0].LessonList;

      newLessonArr.forEach((lesson: any, index: number) => {
        tempActivityResponseData.push({
          [`lesson${index}`]: lesson.Lesson ? lesson.Lesson : null,
          [`startDate${index}`]: new Date(),
          [`endDate${index}`]: new Date(),
          [`developer${index}`]: parseInt(currentUser["Id"]),
          [`dateError${index}`]: false,
        });
      });
      setAtpActivityResponseData([...tempActivityResponseData]);
      setAtpPlannerTemplate(data);
    } else if (key == "data") {
      let lessonArr = [];
      let tempActivityResponseData = [];
      let EditLessonArr = [];
      let tempLessons = data[0].Lessons.split(";");

      tempLessons.forEach((lesson: any) => {
        let lessonSplit = lesson.split("~");
        lessonArr.push({
          ID: lessonSplit[0],
          Lesson: lessonSplit[1],
          Start: lessonSplit[2],
          End: lessonSplit[3],
          Developer: lessonSplit[4],
        });
      });

      lessonArr.forEach((lesson: any, index: number) => {
        tempActivityResponseData.push({
          [`id${index}`]: lesson.ID ? lesson.ID : null,
          [`lesson${index}`]: lesson.Lesson ? lesson.Lesson : null,
          [`startDate${index}`]: lesson.Start ? new Date(lesson.Start) : null,
          [`endDate${index}`]: lesson.End ? new Date(lesson.End) : null,
          [`developer${index}`]: lesson.Developer
            ? parseInt(lesson.Developer)
            : null,
          [`dateError${index}`]: false,
        });
        EditLessonArr.push({
          Lesson: lesson.Lesson ? lesson.Lesson : null,
          StartDate: lesson.Start ? new Date(lesson.Start) : null,
          EndDate: lesson.End ? new Date(lesson.End) : null,
          Developer: lesson.Developer ? lesson.Developer : null,
        });
      });
      setAtpActivityResponseData([...tempActivityResponseData]);
      setAtpLessonData([...EditLessonArr]);
    }
  };
  const atpActivityResponseHandler = (
    index: number,
    key: string,
    value: any
  ) => {
    let newActivityResponseData = [...atpActivityResponseData];
    newActivityResponseData[`${index}`][`${key}`] = value;
    let DataErrorFlag = atpGridDateValidationFunction(
      moment(newActivityResponseData[index][`startDate${index}`]).format(
        "YYYY/MM/DD"
      ),
      moment(newActivityResponseData[index][`endDate${index}`]).format(
        "YYYY/MM/DD"
      )
    );
    newActivityResponseData[`${index}`][`dateError${index}`] = DataErrorFlag;

    setAtpActivityResponseData([...newActivityResponseData]);
  };
  const atpAddActivity = () => {
    let arrBeforeUpdated = [...atpMasterData];
    let LessonJSONArr = [];
    let LessonJSON = "";
    let filtersAtpTemplateDetails = atpTemplatesDetails.data.filter((temp) => {
      return temp.TemplateName == atpPlannerTemplate;
    })[0];

    atpActivityResponseData.forEach((response: any, index: number) => {
      LessonJSONArr.push(
        `${index + 1}~${response[`lesson${index}`]}~${
          response[`startDate${index}`]
        }~${response[`endDate${index}`]}~${response[`developer${index}`]}`
      );
    });
    LessonJSON = LessonJSONArr.join(";");

    // Versions
    let PrjData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type =
        "Project" &&
        arr.Key ==
          (!atpCopy.IsCopy
            ? filtersAtpTemplateDetails.Project
            : atpCopy.Project));
    });
    let PrdData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type =
        "Product" &&
        arr.Key ==
          (!atpCopy.IsCopy
            ? filtersAtpTemplateDetails.Product
            : atpCopy.Product));
    });

    let PrjTitle =
      PrjData.length > 0
        ? PrjData[0].Title
        : (!atpCopy.IsCopy
            ? filtersAtpTemplateDetails.Project
            : atpCopy.Project
          ).replace("V1", "");
    let PrjVersion = PrjData.length > 0 ? PrjData[0].Version : "V1";

    let PrdTitle =
      PrdData.length > 0
        ? PrdData[0].Title
        : (!atpCopy.IsCopy
            ? filtersAtpTemplateDetails.Product
            : atpCopy.Product
          ).replace("V1", "");
    let PrdVersion = PrdData.length > 0 ? PrdData[0].Version : "V1";

    let responseData = {
      Title: `${atpPlannerTemplate} ~ ${findRefID(atpPlannerTemplate)}`,
      Lessons: LessonJSON,
      Types: filtersAtpTemplateDetails.Types,

      Area: filtersAtpTemplateDetails.Area,
      ActivityPlanName: atpCopy.ActivityPlanName,

      ProductCode: filtersAtpTemplateDetails.ProductCode,
      Product: PrdTitle,
      Project: PrjTitle,
      ProductVersion: PrdVersion,
      ProjectVersion: PrjVersion,
      Status: "Scheduled",
      CompletedDate: null,
    };

    sharepointWeb.lists
      .getByTitle(ListName)
      .items.add(responseData)
      .then((item) => {
        arrBeforeUpdated.unshift({
          ID: item.data.Id ? item.data.Id : "",
          TemplateName: responseData.Title ? responseData.Title : "",
          Lessons: responseData.Lessons ? responseData.Lessons : "",
          Project: responseData.Project
            ? responseData.Project + " " + PrjVersion
            : "",
          Types: responseData.Types ? responseData.Types : "",
          ActivityPlanName: atpCopy.ActivityPlanName,
          Area: responseData.Area ? responseData.Area : "",
          Product: responseData.Product
            ? responseData.Product + " " + PrdVersion
            : "",
          ProductCode: responseData.ProductCode ? responseData.ProductCode : "",
          Status: responseData.Status ? responseData.Status : "",
          CompletedDate: responseData.CompletedDate
            ? responseData.CompletedDate
            : null,
        });

        let templateDetails = { ...atpTemplatesDetails };
        templateDetails.templateOptns = [];
        templateDetails.projectOptns = [];
        setAtpTemplatesDetails({ ...templateDetails });
        atpedittemplate = templateDetails;

        setAtpPlannerProduct(null);
        atpGetAllOptions(arrBeforeUpdated, null);
        paginateFunction(1, arrBeforeUpdated);
        setAtpUnsortMasterData([...arrBeforeUpdated]);
        columnSortArr = arrBeforeUpdated;
        setAtpData([...arrBeforeUpdated]);
        columnSortMasterArr = arrBeforeUpdated;
        setAtpMasterData([...arrBeforeUpdated]);
        setAtpFilters({ ...atpFilterKeys });
        setAtpPlannerTemplate(null);
        setAtpAddPlannerPopup(false);
        setAtpActivityResponseData([]);
        setAtpEdit([]);
        setAtpLoader("noLoader");
        ItemAddPopup();

        setatpCopy({
          IsCopy: false,
          Project: "",
          Product: "",
          ActivityPlanName: "",
          IsValidation: false,
        });
      })
      .catch((err) => {
        atpErrorFunction(err, "atpAddActivity");
      });
  };
  const atpUpdateItem = (targetId: number) => {
    let LessonJSONArr = [];
    let LessonJSON = "";
    atpActivityResponseData.forEach((data: any, index: number) => {
      LessonJSONArr.push(
        `${data[`id${index}`]}~${data[`lesson${index}`]}~${
          data[`startDate${index}`]
        }~${data[`endDate${index}`]}~${data[`developer${index}`]}`
      );
    });
    LessonJSON = LessonJSONArr.join(";");

    // Versions
    let PrjData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type = "Project" && arr.Key == atpCopy.Project);
    });
    let PrdData = ProjectOrProductDetails.filter((arr) => {
      return (arr.Type = "Product" && arr.Key == atpCopy.Product);
    });

    let PrjTitle =
      PrjData.length > 0 ? PrjData[0].Title : atpCopy.Project.replace("V1", "");
    let PrjVersion = PrjData.length > 0 ? PrjData[0].Version : "V1";

    let PrdTitle =
      PrdData.length > 0 ? PrdData[0].Title : atpCopy.Product.replace("V1", "");
    let PrdVersion = PrdData.length > 0 ? PrdData[0].Version : "V1";

    let responseData = {
      Lessons: LessonJSON,
      ActivityPlanName: atpCopy.ActivityPlanName,
      Product: PrdTitle,
      Project: PrjTitle,
      ProductVersion: PrdVersion,
      ProjectVersion: PrjVersion,
    };

    sharepointWeb.lists
      .getByTitle(ListName)
      .items.getById(targetId)
      .update(responseData)
      .then(() => {
        let arrBeforeUpdated = [...atpMasterData];
        let targetItem = arrBeforeUpdated.filter((arr) => {
          return arr.ID == atpEdit[0].ID;
        });
        let targetIndex = arrBeforeUpdated.findIndex(
          (arr) => arr.ID == atpEdit[0].ID
        );

        arrBeforeUpdated.splice(targetIndex, 1);

        let updatedTargetItem = {
          ID: targetItem[0].ID,
          TemplateName: targetItem[0].TemplateName,
          Lessons: LessonJSON,
          Project: responseData.Project
            ? responseData.Project + " " + PrjVersion
            : "",
          Types: targetItem[0].Types,
          ActivityPlanName: atpCopy.ActivityPlanName,
          Area: targetItem[0].Area,
          Product: responseData.Product
            ? responseData.Product + " " + PrdVersion
            : "",
          ProductCode: targetItem[0].ProductCode,
          Status: targetItem[0].Status,
          CompletedDate: targetItem[0].CompletedDate,
        };

        arrBeforeUpdated.unshift(updatedTargetItem);

        let filteredItemsAfterUpdated = atpListFilterbyData(arrBeforeUpdated);

        let templateDetails = { ...atpTemplatesDetails };
        templateDetails.templateOptns = [];
        templateDetails.projectOptns = [];
        setAtpTemplatesDetails({ ...templateDetails });
        atpedittemplate = templateDetails;

        setAtpPlannerProduct(null);
        paginateFunction(
          1,
          filteredItemsAfterUpdated.length > 0
            ? [...filteredItemsAfterUpdated]
            : [...arrBeforeUpdated]
        );

        setAtpFilters(
          filteredItemsAfterUpdated.length > 0
            ? { ...atpFilters }
            : { ...atpFilterKeys }
        );
        setAtpUnsortMasterData([...arrBeforeUpdated]);
        columnSortArr =
          filteredItemsAfterUpdated.length > 0
            ? [...filteredItemsAfterUpdated]
            : [...arrBeforeUpdated];
        setAtpData([...columnSortArr]);
        columnSortMasterArr = arrBeforeUpdated;
        setAtpMasterData([...arrBeforeUpdated]);
        setAtpPlannerTemplate(null);
        setAtpEdit([]);
        setAtpActivityResponseData(null);
        setAtpLessonData([]);
        updatePopup();
      })
      .catch((err) => {
        atpErrorFunction(err, "atpUpdateItem");
      });
  };
  const atpDeleteItem = () => {
    sharepointWeb.lists
      .getByTitle(ListName)
      .items.getById(atpDeletePopup.targetID)
      .delete()
      .then(() => {
        let updatedArr = [...atpMasterData];
        let targetIndex = updatedArr.findIndex(
          (arr) => arr.ID == atpDeletePopup.targetID
        );
        updatedArr.splice(targetIndex, 1);

        atpGetAllOptions(updatedArr, null);
        let filteredItemsAfterUpdated = atpListFilterbyData(updatedArr);

        paginateFunction(
          1,
          filteredItemsAfterUpdated.length > 0
            ? [...filteredItemsAfterUpdated]
            : [...updatedArr]
        );

        setAtpFilters(
          filteredItemsAfterUpdated.length > 0
            ? { ...atpFilters }
            : { ...atpFilterKeys }
        );

        setAtpUnsortMasterData([...updatedArr]);
        columnSortArr =
          filteredItemsAfterUpdated.length > 0
            ? [...filteredItemsAfterUpdated]
            : [...updatedArr];
        setAtpData([...columnSortArr]);
        columnSortMasterArr = updatedArr;
        setAtpMasterData([...updatedArr]);
        setAtpDeletePopup({ targetID: null, condition: false });
        setAtpLoader("noLoader");
        DeletePopup();
      })
      .catch((err) => {
        atpErrorFunction(err, "atpDeleteItem");
      });
  };
  const atpGridDateValidationFunction = (startDate: any, EndDate: any) => {
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
  const paginateFunction = (pagenumber, data) => {
    if (data.length > 0) {
      let lastIndex: number = pagenumber * totalPageItems;
      let firstIndex: number = lastIndex - totalPageItems;
      let paginatedItems = data.slice(firstIndex, lastIndex);
      currentpage = pagenumber;
      setAtpDisplayData(paginatedItems);
      setAtpCurrentPage(pagenumber);
    } else {
      setAtpDisplayData([]);
      setAtpCurrentPage(1);
    }
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = atpColumns;
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
    setAtpData([...newDisplayData]);
    setAtpMasterData([...newMasterData]);
    paginateFunction(1, newDisplayData);
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
  const ItemAddPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity plan is successfully submitted !!!")
  );
  const updatePopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity plan is successfully updated !!!")
  );
  const DeletePopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Activity plan is successfully deleted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const atpErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Review log",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        ErrorPopup();
        setAtpLoader("noLoader");
      }
    );
  };
  //Function-Section Ends
  useEffect(() => {
    setAtpLoader("startUpLoader");
    atpGetCurrentUserDetails();
    atpGetTemplates();

    Ap_AnnualPlanId ? getCurrentAPData() : atpGetData();
  }, [atpReRender]);
  return (
    <>
      <div style={{ padding: "5px 10px" }}>
        {/* {atpLoader == "startUpLoader" ? <CustomLoader /> : null} */}
        {atpLoader == "startUpLoader" ? (
          <CustomLoader />
        ) : (
          <>
            {/* Header-Section Starts */}
            <div
              className={styles.atpHeaderSection}
              style={{ paddingBottom: "5px" }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginBottom: 10,
                }}
              >
                {Ap_AnnualPlanId ? (
                  <Icon
                    aria-label="ChevronLeftMed"
                    iconName="NavigateBack"
                    className={atpIconStyleClass.ChevronLeftMed}
                    onClick={() => {
                      props.handleclick("AnnualPlan");
                    }}
                  />
                ) : null}
                <div className={styles.atpHeader}>Activity plan</div>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between" }}>
                <div>
                  <button
                    className={styles.atpAddBtn}
                    onClick={() => {
                      setAtpEdit([]);
                      setAtpLoader("noLoader");
                      setAtpAddPlannerPopup(true);
                      setatpCopy({
                        IsCopy: false,
                        Project: "",
                        Product: "",
                        ActivityPlanName: "",
                        IsValidation: false,
                      });
                    }}
                  >
                    Add Activity planner
                  </button>
                </div>
                {/* {props.isAdmin ? ( */}
                {true ? (
                  <div style={{ display: "flex" }}>
                    <Label
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
                        marginRight: 15,
                      }}
                    >
                      <Icon
                        style={{
                          color: "#1D6F42",
                        }}
                        iconName="ExcelDocument"
                        className={atpIconStyleClass.export}
                      />
                      Export as XLS
                    </Label>
                    <button
                      className={styles.atpAddBtn}
                      onClick={() => {
                        setAtpEdit([]);
                        props.handleclick("ActivityTemplate", null, "AT");
                      }}
                    >
                      Template
                    </button>
                  </div>
                ) : null}
              </div>
              {/* Header-Section Ends */}
              {/* Filter-Section Starts */}
              <div>
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    // justifyContent: "flex-start",
                    justifyContent: "space-between",
                    marginBottom: 5,
                    flexWrap: "wrap",
                  }}
                >
                  <div className={styles.ddSection}>
                    <div>
                      <Label styles={atpLabelStyles}>Type</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          atpFilters.type != "All"
                            ? atpActiveDropdownStyles
                            : atpDropdownStyles
                        }
                        options={atpDropDownOptions.typesOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any) => {
                          atpListFilter("type", option["key"]);
                        }}
                        selectedKey={atpFilters.type}
                      />
                    </div>
                    <div>
                      <Label styles={atpLabelStyles}>Area/Stream</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          atpFilters.area != "All"
                            ? atpActiveDropdownStyles
                            : atpDropdownStyles
                        }
                        options={atpDropDownOptions.areaOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any) => {
                          atpListFilter("area", option["key"]);
                        }}
                        selectedKey={atpFilters.area}
                      />
                    </div>
                    <div>
                      <Label styles={atpLabelStyles}>Product(Program)</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          atpFilters.product != "All"
                            ? atpActiveDropdownStyles
                            : atpDropdownStyles
                        }
                        options={atpDropDownOptions.productOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any) => {
                          atpListFilter("product", option["key"]);
                        }}
                        selectedKey={atpFilters.product}
                      />
                    </div>
                    <div>
                      <Label styles={atpLabelStyles}>Project</Label>
                      <Dropdown
                        placeholder="Select an option"
                        styles={
                          atpFilters.project != "All"
                            ? atpActiveDropdownStyles
                            : atpDropdownStyles
                        }
                        options={atpDropDownOptions.projectOptns}
                        dropdownWidth={"auto"}
                        onChange={(e, option: any) => {
                          atpListFilter("project", option["key"]);
                        }}
                        selectedKey={atpFilters.project}
                      />
                    </div>
                    <div>
                      <Label styles={atpLabelStyles}>Code</Label>
                      <SearchBox
                        styles={
                          atpFilters.code
                            ? atpActiveSearchBoxStyles
                            : atpSearchBoxStyles
                        }
                        value={atpFilters.code}
                        onChange={(e, value) => {
                          atpListFilter("code", value);
                        }}
                      />
                    </div>
                    <div>
                      <Label styles={atpLabelStyles}>Template</Label>
                      <SearchBox
                        styles={
                          atpFilters.template
                            ? atpActiveSearchBoxStyles
                            : atpSearchBoxStyles
                        }
                        value={atpFilters.template}
                        onChange={(e, value) => {
                          atpListFilter("template", value);
                        }}
                      />
                    </div>

                    <div>
                      <Icon
                        iconName="Refresh"
                        title="Click to reset"
                        className={atpIconStyleClass.refresh}
                        onClick={() => {
                          paginateFunction(1, atpUnsortMasterData);
                          columnSortArr = atpMasterData;
                          setAtpData(atpMasterData);
                          columnSortMasterArr = atpMasterData;
                          setAtpMasterData(atpMasterData);
                          setAtpMasterColumns(atpColumns);
                          atpGetAllOptions(atpMasterData, null);
                          setAtpFilters({ ...atpFilterKeys });
                        }}
                      />
                    </div>
                  </div>

                  <div>
                    <Label
                      style={{
                        marginTop: "25px",
                        marginLeft: "10px",
                        fontWeight: "500",
                        color: "#323130",
                        fontSize: "13px",
                      }}
                    >
                      Number of records :{" "}
                      <span style={{ color: "#038387" }}>{atpData.length}</span>
                    </Label>
                  </div>
                </div>
              </div>
            </div>
            {/* Filter-Section Ends */}
            {/* Body-Section Starts */}
            <div>
              {/* DetailList-Section Starts */}
              <div>
                <DetailsList
                  items={atpDisplayData}
                  columns={atpMasterColumns}
                  styles={atpDetailsListStyles}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                />
              </div>
              {atpData.length > 0 ? (
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    margin: "10px 0",
                  }}
                >
                  <Pagination
                    currentPage={atpcurrentPage}
                    totalPages={
                      atpData.length > 0
                        ? Math.ceil(atpData.length / totalPageItems)
                        : 1
                    }
                    onChange={(page) => {
                      paginateFunction(page, atpData);
                    }}
                  />
                </div>
              ) : (
                <Label
                  style={{
                    paddingLeft: 745,
                    paddingTop: 40,
                  }}
                  className={generalStyles.inputLabel}
                >
                  No Data Found !!!
                </Label>
              )}
              {/* DetailList-Section Ends */}
            </div>
            {/* Body-Section Ends */}
            {/* Modal-Section Starts */}
            <div>
              {atpAddPlannerPopup || atpEdit.length > 0 ? (
                <Modal
                  isOpen={
                    atpEdit.length > 0 ? true : false || atpAddPlannerPopup
                  }
                  isBlocking={true}
                  styles={atpModalStyles}
                >
                  <div>
                    <Label className={styles.atpPopupLabel}>
                      {atpEdit.length > 0 && !atpCopy.IsCopy
                        ? "Edit Activity"
                        : "New activity"}
                    </Label>
                    {atpEdit.length == 0 && !atpCopy.IsCopy ? (
                      <div style={{ display: "flex" }}>
                        {/* Product DropDown  */}
                        <div style={{ marginTop: 15 }}>
                          <Label styles={atpModalBoxLabelStyles}>Product</Label>
                          <Dropdown
                            placeholder="Select an option"
                            styles={
                              atpPlannerProduct && atpEdit.length > 0
                                ? atpModalBoxReadOnlyDropdownStyles
                                : atpPlannerProduct
                                ? atpModalBoxActiveDropdownStyles
                                : atpModalBoxDropdownStyles
                            }
                            options={atpTemplatesDetails.productOptns}
                            disabled={atpEdit.length > 0 ? true : false}
                            dropdownWidth={"auto"}
                            onChange={(e, option: any) => {
                              let _projectOptns = [];
                              let templateDetails = { ...atpTemplatesDetails };

                              templateDetails.data.forEach((arr) => {
                                if (
                                  _projectOptns.findIndex(
                                    (_templatesDrpDwn) => {
                                      return (
                                        _templatesDrpDwn.key == arr.Project
                                      );
                                    }
                                  ) == -1 &&
                                  arr.Product == option["key"] &&
                                  arr.Project &&
                                  arr.IsDeleted != true
                                ) {
                                  _projectOptns.push({
                                    key: arr.Project,
                                    text: arr.Project,
                                  });
                                }
                              });
                              templateDetails.projectOptns = [..._projectOptns];
                              setAtpPlannerTemplate(null);
                              setAtpPlannerProject(null);
                              setAtpPlannerProduct(option["key"]);
                              setAtpTemplatesDetails({ ...templateDetails });
                              atpedittemplate = templateDetails;
                            }}
                            selectedKey={atpPlannerProduct}
                          />
                        </div>
                        {/* Project DropDown  */}
                        <div style={{ marginTop: 15 }}>
                          <Label styles={atpModalBoxLabelStyles}>Project</Label>
                          <Dropdown
                            placeholder="Select an option"
                            styles={
                              atpPlannerProject && atpEdit.length > 0
                                ? atpModalBoxReadOnlyDropdownStyles
                                : atpPlannerProject
                                ? atpModalBoxActiveDropdownStyles
                                : atpModalBoxDropdownStyles
                            }
                            options={atpTemplatesDetails.projectOptns}
                            disabled={atpEdit.length > 0 ? true : false}
                            dropdownWidth={"auto"}
                            onChange={(e, option: any) => {
                              let _templatesDrpDwns = [];
                              let templateDetails = { ...atpTemplatesDetails };

                              templateDetails.data.forEach((arr) => {
                                if (
                                  _templatesDrpDwns.findIndex(
                                    (_templatesDrpDwn) => {
                                      return (
                                        _templatesDrpDwn.key == arr.TemplateName
                                      );
                                    }
                                  ) == -1 &&
                                  arr.Project == option["key"] &&
                                  arr.TemplateName &&
                                  arr.IsDeleted != true
                                ) {
                                  _templatesDrpDwns.push({
                                    key: arr.TemplateName,
                                    text: arr.TemplateName,
                                  });
                                }
                              });

                              templateDetails.templateOptns = [
                                ..._templatesDrpDwns,
                              ];
                              setAtpPlannerTemplate(null);
                              setAtpPlannerProject(option["key"]);
                              setAtpTemplatesDetails({ ...templateDetails });
                              atpedittemplate = templateDetails;
                            }}
                            selectedKey={atpPlannerProject}
                          />
                        </div>
                        {/* Template DropDown  */}
                        <div style={{ marginTop: 15 }}>
                          <Label styles={atpModalBoxLabelStyles}>
                            Template
                          </Label>
                          <Dropdown
                            placeholder="Select an option"
                            styles={
                              atpPlannerTemplate && atpEdit.length > 0
                                ? atpModalBoxReadOnlyDropdownStyles
                                : atpPlannerTemplate
                                ? atpModalBoxActiveDropdownStyles
                                : atpModalBoxDropdownStyles
                            }
                            options={atpTemplatesDetails.templateOptns}
                            disabled={atpEdit.length > 0 ? true : false}
                            dropdownWidth={"auto"}
                            onChange={(e, option: any) => {
                              // findRefID(option["key"]);
                              atpActivityResponseArrGenerator(
                                "template",
                                option["key"]
                              );
                            }}
                            selectedKey={atpPlannerTemplate}
                          />
                        </div>
                      </div>
                    ) : (
                      <div
                        style={{
                          width: "300px",
                          paddingRight: !atpCopy.IsCopy ? "35px" : "25px",
                        }}
                      >
                        <Label styles={atpModalBoxLabelStyles}> </Label>
                        <> </>
                      </div>
                    )}
                    {atpPlannerTemplate ? (
                      <div>
                        <div>
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "flex-start",
                              margin: "15px",
                              marginBottom: "10px",
                            }}
                          >
                            <div
                              style={{
                                width: "300px",
                                paddingRight:
                                  !atpCopy.IsCopy && atpEdit.length == 0
                                    ? "35px"
                                    : "25px",
                              }}
                            >
                              <Label>Product</Label>
                              <>
                                {!atpCopy.IsCopy && atpEdit.length == 0 ? (
                                  atpTemplatesDetails.data.filter((temp) => {
                                    return (
                                      temp.TemplateName == atpPlannerTemplate
                                    );
                                  })[0].Product
                                ) : (
                                  <Dropdown
                                    placeholder="Select an option"
                                    styles={atpCopyModalBoxDropdownStyles}
                                    options={atpDropDownOptions.productAllOptns}
                                    dropdownWidth={"auto"}
                                    onChange={(e, option: any) => {
                                      setatpCopy({
                                        IsCopy: atpCopy.IsCopy,
                                        Project: atpCopy.Project,
                                        Product: option["key"],
                                        ActivityPlanName:
                                          atpCopy.ActivityPlanName,
                                        IsValidation: false,
                                      });
                                    }}
                                    selectedKey={atpCopy.Product}
                                  />
                                )}
                              </>
                            </div>
                            <div
                              style={{
                                width: "300px",
                                paddingRight:
                                  !atpCopy.IsCopy && atpEdit.length == 0
                                    ? "35px"
                                    : "25px",
                              }}
                            >
                              <Label>Project</Label>
                              <>
                                {!atpCopy.IsCopy && atpEdit.length == 0 ? (
                                  atpTemplatesDetails.data.filter((temp) => {
                                    return (
                                      temp.TemplateName == atpPlannerTemplate
                                    );
                                  })[0].Project
                                ) : (
                                  <Dropdown
                                    placeholder="Select an option"
                                    styles={atpCopyModalBoxDropdownStyles}
                                    options={atpDropDownOptions.projectAllOptns}
                                    dropdownWidth={"auto"}
                                    onChange={(e, option: any) => {
                                      setatpCopy({
                                        IsCopy: atpCopy.IsCopy,
                                        Project: option["key"],
                                        Product: atpCopy.Product,
                                        ActivityPlanName:
                                          atpCopy.ActivityPlanName,
                                        IsValidation: false,
                                      });
                                    }}
                                    selectedKey={atpCopy.Project}
                                  />
                                )}
                              </>
                            </div>
                            <div style={{ width: "300px" }}>
                              <Label>Activity plan name</Label>
                              <TextField
                                styles={ATTxtBoxStyles}
                                value={atpCopy.ActivityPlanName}
                                // data-index={les.Index}
                                onChange={(e, value: string) => {
                                  setatpCopy({
                                    IsCopy: atpCopy.IsCopy,
                                    Project: atpCopy.Project,
                                    Product: atpCopy.Product,
                                    ActivityPlanName: value,
                                    IsValidation: false,
                                  });
                                }}
                              />
                            </div>
                          </div>
                          <div
                            style={{
                              display: "flex",
                              justifyContent: "flex-start",
                              margin: "15px",
                              marginBottom: "10px",
                            }}
                          >
                            <div
                              style={{ width: "300px", paddingRight: "35px" }}
                            >
                              <Label>Types</Label>
                              <>
                                {atpEdit.length > 0
                                  ? atpEdit[0].Types
                                    ? atpEdit[0].Types
                                    : null
                                  : atpTemplatesDetails.data.filter((temp) => {
                                      return (
                                        temp.TemplateName == atpPlannerTemplate
                                      );
                                    })[0].Types}
                              </>
                            </div>
                            <div
                              style={{ width: "300px", paddingRight: "35px" }}
                            >
                              <Label>Area</Label>
                              <>
                                {atpEdit.length > 0
                                  ? atpEdit[0].Area
                                    ? atpEdit[0].Area
                                    : null
                                  : atpTemplatesDetails.data.filter((temp) => {
                                      return (
                                        temp.TemplateName == atpPlannerTemplate
                                      );
                                    })[0].Area}
                              </>
                            </div>
                            <div
                              style={{ width: "300px", paddingRight: "10px" }}
                            >
                              <Label>Code</Label>
                              <>
                                {atpEdit.length > 0
                                  ? atpEdit[0].ProductCode
                                    ? atpEdit[0].ProductCode
                                    : null
                                  : atpTemplatesDetails.data.filter((temp) => {
                                      return (
                                        temp.TemplateName == atpPlannerTemplate
                                      );
                                    })[0].ProductCode}
                              </>
                            </div>
                          </div>
                        </div>
                        <div
                          style={{
                            marginTop: 30,
                            marginLeft: 15,
                            marginRight: 15,
                            width: 960,
                          }}
                        >
                          {(atpTemplatesDetails.data.filter((temp) => {
                            return temp.TemplateName == atpPlannerTemplate;
                          }).length > 0 &&
                            atpTemplatesDetails.data.filter((temp) => {
                              return temp.TemplateName == atpPlannerTemplate;
                            })[0].LessonList.length > 0) ||
                          atpEdit.length > 0 ? (
                            <DetailsList
                              items={
                                atpEdit.length > 0
                                  ? atpLessonData.length > 0
                                    ? atpLessonData
                                    : []
                                  : atpTemplatesDetails.data.filter((temp) => {
                                      return (
                                        temp.TemplateName == atpPlannerTemplate
                                      );
                                    })[0].LessonList
                              }
                              columns={atpModalBoxColumns}
                              styles={atpModalBoxDetailsListStyles}
                              setKey="set"
                              layoutMode={DetailsListLayoutMode.justified}
                              selectionMode={SelectionMode.none}
                              onShouldVirtualize={() => {
                                return false;
                              }}
                            />
                          ) : (
                            <Label>No Data Found !!!</Label>
                          )}
                        </div>
                      </div>
                    ) : (
                      ""
                    )}
                    <div
                      className={styles.atpModalBoxButtonSection}
                      style={atpPlannerTemplate ? {} : { marginTop: 350 }}
                    >
                      {atpPlannerTemplate &&
                      atpActivityResponseData &&
                      atpActivityResponseData.some(
                        (data, index) => data[`dateError${index}`] == true
                      ) ? (
                        <Label
                          className={generalStyles.dateGridValidationErrorLabel}
                        >
                          *Given end date should not be earlier than the start
                          date
                        </Label>
                      ) : (
                        <>
                          {atpCopy.IsValidation ? (
                            <Label
                              className={
                                generalStyles.dateGridValidationErrorLabel
                              }
                            >
                              *Given data is already exists
                            </Label>
                          ) : (
                            ""
                          )}
                        </>
                      )}
                      {atpPlannerTemplate ? (
                        <button
                          className={
                            atpActivityResponseData &&
                            atpActivityResponseData.some(
                              (data, index) => data[`dateError${index}`] == true
                            )
                              ? styles.atpSubmitBtnDisabled
                              : styles.atpSubmitBtn
                          }
                          onClick={() => {
                            let copyIsValid =
                              atpMasterData.filter((arr) => {
                                return (
                                  arr.Product == atpCopy.Product &&
                                  arr.Project == atpCopy.Project &&
                                  arr.ActivityPlanName ==
                                    atpCopy.ActivityPlanName
                                );
                              }).length > 0
                                ? true
                                : false;

                            if (
                              atpActivityResponseData &&
                              !atpActivityResponseData.some(
                                (data, index) =>
                                  data[`dateError${index}`] == true
                              )
                            ) {
                              if (atpEdit.length > 0 && !atpCopy.IsCopy) {
                                setAtpLoader("ActivitySubmitLoader");
                                atpUpdateItem(atpEdit[0].ID);
                              } else {
                                if (atpPlannerTemplate) {
                                  // !atpCopy.IsValidation && !copyIsValid
                                  true
                                    ? (setAtpLoader("ActivitySubmitLoader"),
                                      atpAddActivity())
                                    : setatpCopy({
                                        IsCopy: true,
                                        Project: atpCopy.Project,
                                        Product: atpCopy.Product,
                                        ActivityPlanName:
                                          atpCopy.ActivityPlanName,
                                        IsValidation: true,
                                      });
                                }
                              }
                            }

                            if (
                              atpLoader == "ActivitySubmitLoader" &&
                              atpEdit.length > 0
                            ) {
                              // Non dropdown value remove
                              let isOriginalData_Project =
                                ProjectOrProductDetails.filter((arr) => {
                                  return (arr.Type =
                                    "Project" && arr.Key == atpEdit[0].Project);
                                });
                              let isOriginalData_Product =
                                ProjectOrProductDetails.filter((arr) => {
                                  return (arr.Type =
                                    "Product" && arr.Key == atpEdit[0].Product);
                                });

                              if (
                                isOriginalData_Product.length == 0 &&
                                atpEdit[0].Product
                              ) {
                                atpDropDownOptions.productAllOptns.pop();
                              }
                              if (
                                isOriginalData_Project.length == 0 &&
                                atpEdit[0].Project
                              ) {
                                atpDropDownOptions.projectAllOptns.pop();
                              }
                              setAtpDropDownOptions({ ...atpDropDownOptions });
                            }
                          }}
                        >
                          <span>
                            {atpLoader == "ActivitySubmitLoader" ? (
                              <Spinner />
                            ) : (
                              <>
                                <Icon
                                  iconName="Save"
                                  style={{
                                    position: "relative",
                                    top: 3,
                                    left: -8,
                                  }}
                                />
                                {atpEdit.length > 0 && !atpCopy.IsCopy
                                  ? "Update"
                                  : "Submit"}
                              </>
                            )}
                          </span>
                        </button>
                      ) : (
                        ""
                      )}
                      <button
                        className={styles.atpCloseBtn}
                        onClick={() => {
                          setatpCopy({
                            IsCopy: false,
                            Project: "",
                            Product: "",
                            ActivityPlanName: "",
                            IsValidation: false,
                          });

                          // Non dropdown value remove
                          if (atpEdit.length > 0) {
                            let isOriginalData_Project =
                              ProjectOrProductDetails.filter((arr) => {
                                return (arr.Type =
                                  "Project" && arr.Key == atpEdit[0].Project);
                              });
                            let isOriginalData_Product =
                              ProjectOrProductDetails.filter((arr) => {
                                return (arr.Type =
                                  "Product" && arr.Key == atpEdit[0].Product);
                              });

                            if (
                              isOriginalData_Product.length == 0 &&
                              atpEdit[0].Product
                            ) {
                              atpDropDownOptions.productAllOptns.pop();
                            }
                            if (
                              isOriginalData_Project.length == 0 &&
                              atpEdit[0].Project
                            ) {
                              atpDropDownOptions.projectAllOptns.pop();
                            }
                            setAtpDropDownOptions({ ...atpDropDownOptions });
                          }

                          if (atpEdit.length > 0) {
                            let templateDetails = { ...atpTemplatesDetails };
                            templateDetails.templateOptns = [];
                            templateDetails.projectOptns = [];
                            setAtpTemplatesDetails({ ...templateDetails });
                            atpedittemplate = templateDetails;
                            setAtpEdit([]);
                            setAtpLessonData([]);
                            setAtpPlannerTemplate(null);
                            setAtpPlannerProduct(null);
                            setAtpPlannerProject(null);
                            setAtpActivityResponseData([]);
                          } else {
                            let templateDetails = { ...atpTemplatesDetails };
                            templateDetails.templateOptns = [];
                            templateDetails.projectOptns = [];
                            setAtpTemplatesDetails({ ...templateDetails });
                            atpedittemplate = templateDetails;
                            setAtpAddPlannerPopup(false);
                            setAtpPlannerTemplate(null);
                            setAtpPlannerProduct(null);
                            setAtpPlannerProject(null);
                            setAtpActivityResponseData([]);
                          }
                        }}
                      >
                        <span>
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
              ) : (
                ""
              )}
            </div>
            <div>
              {atpDeletePopup.condition ? (
                <Modal isOpen={atpDeletePopup.condition} isBlocking={true}>
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
                        Delete Activity plan
                      </Label>
                      <Label className={styles.deletePopupDesc}>
                        Are you sure you want to delete activity plan?
                      </Label>
                    </div>
                  </div>
                  <div className={styles.apDeletePopupBtnSection}>
                    <button
                      onClick={(_) => {
                        setAtpLoader("DeleteLoader");
                        atpDeleteItem();
                      }}
                      className={styles.apDeletePopupYesBtn}
                    >
                      {atpLoader == "DeleteLoader" ? <Spinner /> : "Yes"}
                    </button>
                    <button
                      onClick={(_) => {
                        setAtpDeletePopup({ targetID: null, condition: false });
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
            {/* Modal-Section Ends */}
          </>
        )}
      </div>
    </>
  );
};

export default ActivityPlan;
