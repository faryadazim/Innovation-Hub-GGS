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
    IDropdownOption,
    NormalPeoplePicker,
    Persona,
    PersonaPresence,
    PersonaSize,
    Modal,
    DatePicker,
    IDatePickerStyles,
    Checkbox,
    ICheckboxStyles,
    TextField,
    ITextFieldStyles,
    Spinner,
    TooltipHost,
    TooltipDelay,
    TooltipOverflowMode,
    DirectionalHint,
    IColumn, Stack, StackItem, PrimaryButton, IButtonStyles
} from "@fluentui/react";



import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import Pagination from "office-ui-fabric-react-pagination";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";
import { sp } from "@pnp/sp";
import { Log } from "@microsoft/sp-core-library";
import { Async } from "office-ui-fabric-react";
import { xor } from "lodash";

let columnSortArr = [];
let columnSortMasterArr = [];
let editYear = [];
let DateListFormat = "DD/MM/YYYY";

const Dashboard = (props: any) => {


    const today = moment();
    const weekNumber = today.format('w');


    const sharepointWeb = Web(props.URL);
    const ListNameURL = props.WeblistURL;
    const pbDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
        },
        title: {
            backgroundColor: "#F5F5F7",
            fontSize: 12,
            color: "#7C7C7C",
            border: "1px solid #E8E8EA",
            borderRadius: "4px",
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
    const pbActiveDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
        },
        title: {
            backgroundColor: "#F5F5F7",
            fontSize: 12,
            color: "#038387",
            border: "2px solid #038387",
            borderRadius: "4px",
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

    let Ap_AnnualPlanId = props.AnnualPlanId;
    const headingStyles: Partial<ILabelStyles> = {
        root: {
            color: "#000",
            fontSize: 26,
            padding: 0,
            marginBottom: 10,
        },
    };
    let loggeduseremail: string = props.spcontext.pageContext.user.email;

    let currentpage = 1;
    let totalPageItems = 10;
    const apAllitems = [];
    const ReviewLogAllItems = [];
    const apMasterProductCollection = [];
    const allPeoples = props.peopleList;
    console.log(allPeoples, "allPeoples");

    const BAacronymsCollection = [
        {
            Name: "PD Curriculum",
            ShortName: "PDC",
        },
        {
            Name: "PD Professional Learning",
            ShortName: "PDPL",
        },
        {
            Name: "PD School Improvements",
            ShortName: "PDSI",
        },
        {
            Name: "SS Business",
            ShortName: "SSB",
        },
        {
            Name: "SS Publishing",
            ShortName: "SSP",
        },
        {
            Name: "SS Content Creation",
            ShortName: "SSCC",
        },
        {
            Name: "SS Marketing",
            ShortName: "SSM",
        },
        {
            Name: "SS Technology",
            ShortName: "SST",
        },
        {
            Name: "SS Research and Evaluation",
            ShortName: "SSRE",
        },
        {
            Name: "SD School Partnerships",
            ShortName: "SSPSP",
        },
    ];
    // const TODacronymsCollection = [
    //   {
    //     Name: "Product",
    //     ShortName: "PT",
    //   },
    //   {
    //     Name: "Project",
    //     ShortName: "PR",
    //   },
    //   {
    //     Name: "Task",
    //     ShortName: "T",
    //   },
    //   {
    //     Name: "Activity",
    //     ShortName: "A",
    //   },
    //   {
    //     Name: "Product initiative",
    //     ShortName: "PTI",
    //   },
    //   {
    //     Name: "Product tool",
    //     ShortName: "PRT",
    //   },
    //   {
    //     Name: "Product related",
    //     ShortName: "TPR",
    //   },
    //   {
    //     Name: "Activity planner",
    //     ShortName: "AP",
    //   },
    //   {
    //     Name: "Organisation solution",
    //     ShortName: "SPT",
    //   },
    //   {
    //     Name: "Project solution",
    //     ShortName: "SPR",
    //   },
    //   {
    //     Name: "Task solution",
    //     ShortName: "ST",
    //   },
    //   {
    //     Name: "Activity solution",
    //     ShortName: "BA",Das
    //   },
    //   {
    //     Name: "Test",
    //     ShortName: "TTF",
    //   },
    //   {
    //     Name: "Product initiative",
    //     ShortName: "NI",
    //   },
    //   {
    //     Name: "Product tool",
    //     ShortName: "TEC",
    //   },
    //   {
    //     Name: "Product related",
    //     ShortName: "S",
    //   },
    //   {
    //     Name: "Organisation solution",
    //     ShortName: "OS",
    //   },
    //   {
    //     Name: "Project solution",
    //     ShortName: "SS",
    //   },
    //   {
    //     Name: "Task solution",
    //     ShortName: "TS",
    //   },
    //   {
    //     Name: "Activity solution",
    //     ShortName: "AS",
    //   },
    // ];

    const TODacronymsCollection = [
        {
            Name: "Product",
            ShortName: "PT",
        },
        {
            Name: "Project",
            ShortName: "PR",
        },
        {
            Name: "Task",
            ShortName: "T",
        },
        {
            Name: "Activity",
            ShortName: "A",
        },
        {
            Name: "Product initiative",
            ShortName: "NI",
        },
        {
            Name: "Technology",
            ShortName: "TEC",
        },
        {
            Name: "Strategy",
            ShortName: "S",
        },
        {
            Name: "Activity planner",
            ShortName: "AP",
        },
        {
            Name: "Organisation solution",
            ShortName: "OS",
        },
        {
            Name: "System solution",
            ShortName: "SS",
        },
        {
            Name: "Task solution",
            ShortName: "TS",
        },
        {
            Name: "Activity solution",
            ShortName: "AS",
        },
        {
            Name: "Test",
            ShortName: "TTF",
        },
    ];

    const apColumns = [

        {
            key: "Product",
            name: "Product or Solution",
            fieldName: "Product",
            minWidth: 130,
            maxWidth: 280,
            // onColumnClick: (
            //     ev: React.MouseEvent<HTMLElement>,
            //     column: IColumn
            // ) => {
            //     _onColumnClick(ev, column);
            // },
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
            key: "ProjectOrTask",
            name: "Name of the deliverable",
            fieldName: "ProjectOrTask",
            minWidth: 210,
            maxWidth: 220,
            // onColumnClick: (
            //     ev: React.MouseEvent<HTMLElement>,
            //     column: IColumn
            // ) => {
            //     _onColumnClick(ev, column);
            // },
            onRender: (item) => (
                <>
                    <TooltipHost
                        id={item.ID}
                        content={item.ProjectOrTask + " " + item.ProjectVersion}
                        overflowMode={TooltipOverflowMode.Parent}
                    >
                        <span aria-describedby={item.ID}>
                            {item.ProjectOrTask + " " + item.ProjectVersion}
                        </span>
                    </TooltipHost>
                </>
            ),
        },

        {
            key: "AP",
            name: "AP",
            fieldName: "AP",
            minWidth: 30,
            maxWidth: 70,

            onRender: (item) => (
                <>
                    <Icon
                        style={{
                            marginLeft: 0,
                        }}
                        iconName="Link12"
                        className={apIconStyleClass.link}
                        onClick={() => {
                            props.handleclick("ActivityPlan", item.ID);
                        }}
                    />
                </>
            ),
        },
        {
            key: "DP/AP",
            name: "DP",
            fieldName: "DPAP",
            minWidth: 30,
            maxWidth: 70,

            onRender: (item) => (
                <>
                    <Icon
                        style={{
                            marginLeft: 0,
                        }}
                        iconName="Link12"
                        className={apIconStyleClass.link}
                        onClick={() => {
                            props.handleclick("DeliveryPlan", item.ID, "AP");
                        }}
                    />
                </>
            ),
        },

        {
            key: "Status",
            name: "Status",
            fieldName: "Status",
            minWidth: 130,
            maxWidth: 130,
            // onColumnClick: (
            //     ev: React.MouseEvent<HTMLElement>,
            //     column: IColumn
            // ) => {
            //     _onColumnClick(ev, column);
            // },
            onRender: (item) => (
                <>
                    {item.Status == "Completed" ? (
                        <div className={apStatusStyleClass.completed}>
                            {item.Status}
                        </div>
                    ) : item.Status == "Scheduled" ? (
                        <div className={apStatusStyleClass.scheduled}>
                            {item.Status}
                        </div>
                    ) : item.Status == "On schedule" ? (
                        <div className={apStatusStyleClass.onSchedule}>
                            {item.Status}
                        </div>
                    ) : item.Status == "Behind schedule" ? (
                        <div className={apStatusStyleClass.behindScheduled}>
                            {item.Status}
                        </div>
                    ) : item.Status == "On hold" ? (
                        <div className={apStatusStyleClass.Onhold}>{item.Status}</div>
                    ) : (
                        ""
                    )}
                </>
            ),
        },

    ];

    const apDrpDwnOptns = {
        baOptns: [{ key: "All", text: "All" }],
        todOptns: [{ key: "All", text: "All" }],
        potOptns: [{ key: "All", text: "All" }],
        managerOptns: [{ key: "All", text: "All" }],
        PriorityOptns: [{ key: "All", text: "All" }],
        developerOptns: [{ key: "All", text: "All" }],
        termOptns: [{ key: "All", text: "All" }],
        yearOptns: [],
        weekOptns: []
    };
    const apModalBoxDrpDwnOptns = {
        baOptns: [],
        todOptns: [],
        potOptns: [],
        managerOptns: [],
        developerOptns: [],
        PriorityOptns: [],
        termOptns: [],
        productOptns: [],
        yearOptns: [],
        statusOtpns: [
            { key: "On hold", text: "On hold" },
            { key: "Completed", text: "Completed" },
        ],
    };
    // const apFilterKeys = {
    //     ProjectOrTaskSearch: "",
    //     BusinessArea: "All",
    //     TypeOfProject: "All",
    //     ProjectOrTask: "All",
    //     PM: "All",
    //     D: "All",
    //     Term: "All",
    //     Year: "2023",
    //     Product: "All",
    //     CurrentMonth: "All",
    //     CurrentWeek: "All"
    // };
    const responseData = {
        ID: null,
        businessArea: "",
        typeOfProject: "",
        term: [],
        product: "",
        startDate: new Date(),
        endDate: new Date(),
        projectOrTask: "",
        Priority: "",
        ProjectVersion: "",
        year: "",
        manager: "",
        developer: [],
        status: "",
    };
    const apErrorStatus = {
        businessAreaError: "",
        projectOrTaskError: "",
        productError: "",
    };

    //StylesStart

    const primaryColor = '#177F99';

    const buttonStyles: IButtonStyles = {
        root: {
            borderRadius: 0,
            backgroundColor: 'white',
            borderColor: primaryColor,
            borderWidth: '2px',
            marginLeft: '0px',
            marginRight: '0px',
            color: primaryColor,
            minWidth: 100,
            padding: '10px 20px',
            selectors: {
                ':hover': {
                    backgroundColor: primaryColor,
                    color: 'white',
                },
            },
        },
    };

    const buttonStylesBorderLess: IButtonStyles = {
        root: {
            borderRadius: 0,
            backgroundColor: 'white',
            borderColor: primaryColor,
            borderWidth: '2px',
            marginLeft: '0px',
            borderLeft: '0px',
            borderRight: '0px',
            marginRight: '0px',
            color: primaryColor,
            minWidth: 100,
            padding: '10px 20px',
            selectors: {
                ':hover': {
                    backgroundColor: primaryColor,
                    color: 'white',
                },
            },
        },
    };

    const activeButtonStyles: IButtonStyles = {
        root: {
            borderRadius: 0,
            backgroundColor: primaryColor,
            marginLeft: '0px',
            marginRight: '0px',
            borderColor: primaryColor,
            borderWidth: '2px',
            color: 'white',
            minWidth: 100,
            padding: '10px 20px',
            selectors: {
                ':hover': {
                    backgroundColor: primaryColor,
                    color: 'white',
                },
            },
        },
    };












    const elementStyle = {
        padding: '2px',
        paddingLeft: '4px',
        backgroundColor: '#177F99',
        marginBottom: '0px',
        marginTop: '2px',
        paddingTop: '2px',
        paddingBottom: '2px',
        color: 'white',
    };
    const drDetailsListStyles: Partial<IDetailsListStyles> = {
        root: {
            width: '100%',
            selectors: {
                ".ms-DetailsRow-cell": {
                    height: 15,
                },
            },
        },
        contentWrapper: {
            height: "calc(100vh - 175px)",
            overflowX: "hidden",
            overflowY: "auto",
        },
    };
    const drDetailsListStyleSmallerThreeCount: Partial<IDetailsListStyles> = {
        root: {
            width: '100%',
            selectors: {
                ".ms-DetailsRow-cell": {
                    height: 15,
                },
            },
        },
        contentWrapper: {
            height: "90px",
            overflowX: "hidden",
            overflowY: "auto",
        },
    };
    const gridStyles: Partial<IDetailsListStyles> = {
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
    const apLabelStyles: Partial<ILabelStyles> = {
        root: {
            width: 150,
            marginRight: 10,
            fontSize: "13px",
            color: "#323130",
        },
    };
    const apShortLabelStyles: Partial<ILabelStyles> = {
        root: {
            width: 75,
            marginRight: 10,
            fontSize: "13px",
            color: "#323130",
        },
    };

    const DBfilterShortLabelStyles: Partial<ILabelStyles> = {
        root: {
            width: 75,
            marginRight: 10,
            fontSize: "13px",
            color: "#323130",
        },
    };

    const pbLabelStyles: Partial<ILabelStyles> = {
        root: {
            width: 150,
            marginRight: 10,
            fontSize: "13px",
            color: "#323130",
        },
    };
    const apSearchBoxStyles: Partial<ISearchBoxStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
            outline: "none",
            border: "1px solid #E8E8EA",
            borderRadius: "4px",
        },
        icon: { fontSize: 14, color: "#000" },
    };
    const apActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
            outline: "none",
            color: "#038387",
            border: "2px solid #038387",
            borderRadius: "4px",
        },
        icon: { fontSize: 14, color: "#038387" },
    };
    const apDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
        },
        title: {
            backgroundColor: "#F5F5F7",
            fontSize: 12,
            color: "#7C7C7C",
            border: "1px solid #E8E8EA",
            borderRadius: "4px",
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
            maxHeight: "400px !important",
        },
        caretDown: { fontSize: 14, color: "#000" },
    };
    const apActiveDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 186,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
        },
        title: {
            backgroundColor: "#F5F5F7",
            fontSize: 12,
            color: "#038387",
            border: "2px solid #038387",
            borderRadius: "4px",
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
            maxHeight: "400px !important",
        },
        caretDown: { fontSize: 14, color: "#000" },
    };
    const apShortDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 75,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
        },
        title: {
            backgroundColor: "#F5F5F7",
            fontSize: 12,
            color: "#7C7C7C",
            border: "1px solid #E8E8EA",
            borderRadius: "4px",
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
    const apActiveShortDropdownStyles: Partial<IDropdownStyles> = {
        root: {
            width: 75,
            marginRight: "15px",
            backgroundColor: "#F5F5F7",
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
        caretDown: { fontSize: 14, color: "#000" },
        callout: {
            maxHeight: 300,
        },
    };
    const apModalBoxDropdownStyles: Partial<IDropdownStyles> = {
        root: { width: "300px", margin: "10px 20px" },
        title: {
            fontSize: 12,
            borderRadius: "4px",
            border: "1px solid",
            height: "36px",
            padding: "3px 10px",
            color: "#000",
        },
        dropdownItemsWrapper: { fontSize: 12 },
        dropdownItemSelected: { fontSize: 12 },
        caretDown: {
            fontSize: 14,
            padding: "3px",
            color: "#000",
            fontWeight: "bold",
        },
    };
    const apModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
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
    const apTxtBoxStyles: Partial<ITextFieldStyles> = {
        root: { width: "780px", margin: "10px 20px", borderRadius: "4px" },
        field: { fontSize: 12, color: "#000" },
    };
    const apTxtBoxStylesSmall: Partial<ITextFieldStyles> = {
        root: { width: "100px", margin: "10px 10px", borderRadius: "4px" },
        field: { fontSize: 12, color: "#000" },
    };
    const apModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
        root: {
            width: "300px",
            margin: "10px 20px",
            borderRadius: "4px",
        },
        icon: {
            fontSize: "17px",
            color: "#000",
            fontWeight: "bold",
        },
    };
    const apModalBoxCheckBoxStyles: Partial<ICheckboxStyles> = {
        root: { marginTop: "46px", transform: "translateX(-26px)" },
        label: { fontWeight: "600" },
    };
    const apModalBoxPP = mergeStyles({
        width: "300px",
        margin: "10px 20px",
    });
    const apIconStyle = mergeStyles({
        fontSize: 17,
        height: 14,
        width: 17,
        cursor: "pointer",
    });
    const apIconStyleClass = mergeStyleSets({
        link: [{ color: "#2392B2", margin: "0" }, apIconStyle],
        delete: [{ color: "#CB1E06", margin: "0 7px " }, apIconStyle],
        edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, apIconStyle],
        refresh: [
            {
                color: "white",
                fontSize: "18px",
                height: 20,
                width: 22,
                cursor: "pointer",
                backgroundColor: "#038387",
                padding: 5,
                marginTop: 27,
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
    const apStatusStyle = mergeStyles({
        textAlign: "center",
        borderRadius: "25px",
        width: '108px'
    });
    const apStatusStyleClass = mergeStyleSets({
        completed: [
            {
                fontWeight: "600",
                padding: "3px",
                color: "#438700",
                backgroundColor: "#D9FFB3",
            },
            apStatusStyle,
        ],
        scheduled: [
            {
                fontWeight: "600",
                padding: "3px",
                color: "#9f6700",
                backgroundColor: "#FFDB99",
            },
            apStatusStyle,
        ],
        onSchedule: [
            {
                fontWeight: "600",
                padding: "3px",
                color: "#B3B300 ",
                backgroundColor: "#FFFFB3",
            },
            apStatusStyle,
        ],
        behindScheduled: [
            {
                fontWeight: "600",
                padding: "3px",
                color: "#FF0000",
                backgroundColor: "#FFB3B3",
            },
            apStatusStyle,
        ],
        Onhold: [
            {
                fontWeight: "600",
                padding: "3px",
                color: "#773030",
                backgroundColor: "#e6b1b1",
            },
            apStatusStyle,
        ],
    });

    //stylesEnd

    const DRAColumns: IColumn[] = [
        {
            key: "column1",
            name: "Deliverable",
            fieldName: "Diliverable",
            minWidth: 100,
            maxWidth: 250,
            onRender: (item) => (
                <a href={item.DiliverableLink} target="_blank">
                    {item.Diliverable}
                </a>
            )
            // onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
            //   _onColumnClick(ev, column);
            // },
        },
        {
            key: "column2",
            name: "Developer",
            fieldName: "Devoloper",
            minWidth: 100,
            maxWidth: 250,
            onRender: (item) => (
                <div style={{ display: "flex" }}>
                    <div
                        style={{
                            marginTop: "-6px",
                        }}
                        title={item.UserName}
                    >
                        <Persona
                            size={PersonaSize.size32}
                            presence={PersonaPresence.none}
                            imageUrl={
                                "/_layouts/15/userphoto.aspx?size=S&username=" +
                                `${item.UserEmail}`
                            }
                        />
                    </div>
                    <div>
                        <span title={item.UserName} style={{ fontSize: "13px" }}>
                            {item.Devoloper}
                        </span>
                    </div>
                </div>
            )
            // onColumnClick: (ev: React.MouseEvent<HTMLElement>, column: IColumn) => {
            //   _onColumnClick(ev, column);
            // },
        },

    ];

    let Pb_Year = moment().year();
    // let Pb_NextWeekYear = moment().add(1, "week").year();
    // let Pb_LastWeekYear = moment().subtract(1, "week").year();
    //let Pb_WeekNumber = moment().isoWeek();
    const apFilterKeys = {
        ProjectOrTaskSearch: "",
        BusinessArea: "All",
        TypeOfProject: "All",
        ProjectOrTask: "All",
        PM: "All",
        D: "All",
        Term: "All",
        Year: Pb_Year,
        Week: "All",
        Product: "All",
        CurrentMonth: "All",
        CurrentWeek: "All",
        CurrentDay: "All",

    };
    const totalWeekOfYear: {
        key: any,
        text: any,
    }[] = [{ key: 'All', text: 'All' },
    { key: 1, text: 1 },
    { key: 2, text: 2 },
    { key: 3, text: 3 },
    { key: 4, text: 4 },
    { key: 5, text: 5 },
    { key: 6, text: 6 },
    { key: 7, text: 7 },
    { key: 8, text: 8 },
    { key: 9, text: 9 },
    { key: 10, text: 10 },
    { key: 11, text: 11 },
    { key: 12, text: 12 },
    { key: 13, text: 13 },
    { key: 14, text: 14 },
    { key: 15, text: 15 },
    { key: 16, text: 16 },
    { key: 17, text: 17 },
    { key: 18, text: 18 },
    { key: 19, text: 19 },
    { key: 20, text: 20 },
    { key: 21, text: 21 },
    { key: 22, text: 22 },
    { key: 23, text: 23 },
    { key: 24, text: 24 },
    { key: 25, text: 25 },
    { key: 26, text: 26 },
    { key: 27, text: 27 },
    { key: 28, text: 28 },
    { key: 29, text: 29 },
    { key: 30, text: 30 },
    { key: 31, text: 31 },
    { key: 32, text: 32 },
    { key: 33, text: 33 },
    { key: 34, text: 34 },
    { key: 35, text: 35 },
    { key: 36, text: 36 },
    { key: 37, text: 37 },
    { key: 38, text: 38 },
    { key: 39, text: 39 },
    { key: 40, text: 40 },
    { key: 41, text: 41 },
    { key: 42, text: 42 },
    { key: 43, text: 43 },
    { key: 44, text: 44 },
    { key: 45, text: 45 },
    { key: 46, text: 46 },
    { key: 47, text: 47 },
    { key: 48, text: 48 },
    { key: 49, text: 49 },
    { key: 50, text: 50 },
    { key: 51, text: 51 },
    { key: 52, text: 52 }
        ]


    const [ProductDropDownoptions, setProductDropDownoptions] = useState<IDropdownOption[]>([{ key: 'All', text: 'All' },]);
    const [currentUser, setcurrentUser] = useState<any>({});
    const [BussinessDropDownoptions, setBussinessDropDownoptions] = useState<IDropdownOption[]>([{ key: 'All', text: 'All' },])
    const [pbFilterData, setpbFilterData] = useState([]);
    const [apReRender, setapReRender] = useState(true);
    const [stackProps, setStackProps] = useState({ horizontal: true });
    const [stackItemWidth, setstackItemWidth] = useState({
        leftSec: "60%",
        rightSection: "60%",
        mobileView: false
    })

    const [ReRenderState, setReRenderState] = useState(false);
    const [apUnsortMasterData, setApUnsortMasterData] = useState(apAllitems);
    const [apMasterData, setApMasterData] = useState(apAllitems);
    const [apData, setApData] = useState(apAllitems);
    const [displayData, setdisplayData] = useState(apAllitems);
    const [yearDropDown, setYearDropDown] = useState([]);
    const [weekDropDowns, setWeekDropDowns] = useState(totalWeekOfYear);

    const [reviewLogUnsortedData, setreviewLogUnsortedData] = useState(ReviewLogAllItems)
    const [reviewLogSortedData, setreviewLogSortedData] = useState(ReviewLogAllItems)
    const [reviewLogSortedDataToBeDisplayed, setreviewLogSortedDataToBeDisplayed] = useState(ReviewLogAllItems)
    const [apResponseData, setApResponseData] = useState(responseData);
    const [apMasterProducts, setApMasterProducts] = useState(
        apMasterProductCollection
    );
    const [apDropDownOptions, setApDropDownOptions] = useState(apDrpDwnOptns);
    const [apModalBoxDropDownOptions, setApModalBoxDropDownOptions] = useState(
        apModalBoxDrpDwnOptns
    );


    const [reRenderApplicationData, setreRenderApplicationData] = useState(false)
    const [apFilterOptions, setApFilterOptions] = useState(apFilterKeys);
    const [apModalBoxVisibility, setApModalBoxVisibility] = useState({
        condition: false,
        action: "",
        selectedItem: [],
    });
    const [apDeletePopup, setApDeletePopup] = useState({
        condition: false,
        targetId: 0,
    });
    const [submitConfirmationPopup, setSubmitConfirmationPopup] = useState(false);
    const [apModelBoxDrpDwnToTxtBox, setApModelBoxDrpDwnToTxtBox] =
        useState(false);
    const [apcurrentPage, setApCurrentPage] = useState(currentpage);
    const [apShowMessage, setApShowMessage] = useState(apErrorStatus);
    const [apStartUpLoader, setApStartUpLoader] = useState(true);
    const [apOnSubmitLoader, setApOnSubmitLoader] = useState(false);
    const [apOnDeleteLoader, setApOnDeleteLoader] = useState(false);
    const [apSubmitConfirmLoader, setApSubmitConfirmLoader] = useState(false);
    const [masterApColumn, setMasterApColumn] = useState(apColumns);
    const [DRAColumn, SetDRAColumns] = useState<IColumn[]>(DRAColumns);
    const [historyData, setHistory] = useState({
        condition: false,
        sourcePage: "",
        targetID: null,
    });
    const [userOnlyDataForapAllitems, setUserOnlyDataForapAllitems] = useState([]);
    const [allUserDataForapAllitems, setAllUserDataForapAllitems] = useState([]);

    const [pageSwitch, setPageSwitch] = useState("User");
    const [allUserList, setAllUserList] = useState([]);
    const getApData = async () => {
        var listOfUniqueAnualIds: string[] = [];
        await sharepointWeb.lists
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
                "Master_x0020_Project/ProductVersion",
                "FieldValuesAsText/StartDate",
                "FieldValuesAsText/Modified",
                "FieldValuesAsText/PlannedEndDate"
            )
            .expand(
                "ProjectOwner",
                "ProjectLead",
                "Master_x0020_Project",
                "FieldValuesAsText"
            )
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then((items) => {
                console.log(items , "core data")
                let tempArrCustom = [];
                items.forEach((item: any, index: number) => {
                    const returnFiltered = allitemsArrayFormatter(item, apAllitems);
                    tempArrCustom.push(returnFiltered)

                });
                let productOption: IDropdownOption[] = [];

                const filteredProducts = tempArrCustom.reduce((acc, cur) => {
                    if (cur.value !== '' && !acc.some((item) => item.Product === cur.Product)) {
                        acc.push(cur);
                    }
                    return acc;
                }, []);
                filteredProducts.map((x) => {
                    productOption.push({ key: x.Product, text: x.Product })
                })
                //findning unique values for product 
                setProductDropDownoptions([...ProductDropDownoptions, ...productOption])
                filterKeys(items);




                // finding unique values for user 
                tempArrCustom.map((zz) => {
                    zz?.listOfDevIds.map((x) => {
                        if (
                            listOfUniqueAnualIds.findIndex((baOptn) => {
                                return baOptn == x;
                            }) == -1
                        ) {
                            listOfUniqueAnualIds.push(x)
                        }
                    })

                }),


                    console.log(listOfUniqueAnualIds, "uniqueAnualIds")
                //storing all users data for other tabs use 
                setAllUserDataForapAllitems([...tempArrCustom]);

                let tempArr = [...tempArrCustom];
                let tempApFilterKeys = { ...apFilterOptions };
                sp.web.currentUser.get().then(user => {
                    setcurrentUser(user)
                    console.log(user, "Current login user")
                    let devArr = [];
                    tempArr.forEach((arr) => {
                        if (arr.DNames.length != 0) {
                            if (arr.DNames.some((DName) => DName.name == user.Title)) {
                                devArr.push(arr);
                            }
                        }
                    });

                    //login user specifc data 
                    setUserOnlyDataForapAllitems([...devArr])

                    setApUnsortMasterData([...devArr]);
                    columnSortArr = [...devArr];
                    setApData([...devArr]);
                    columnSortMasterArr = [...devArr];
                    setApMasterData([...devArr]);
                    paginate(1);
                    columnSortArr = [...devArr];
                    setApData([...devArr]);
                    columnSortMasterArr = [...devArr];
                    setApMasterData([...devArr]);
                    setMasterApColumn(apColumns);
                    filterKeysAfterModified(apMasterData);
                    paginatewithdata(1, [...devArr]);
                    setApFilterOptions({ ...apFilterKeys });

                    const currentYearData = devArr.filter((arr) => {
                        return arr.Year == Pb_Year;
                    });
                    let tempArrTestDrive = [...currentYearData];
                    paginatewithdata(1, [...tempArrTestDrive]);
                    setApData(tempArrTestDrive)

                });



                // setApStartUpLoader(false);
                // const pageName = new URLSearchParams(window.location.search).get("TOD");
                // pageName ? checkForTOD(pageName) : null;



                // setdisplayData(apData);
            })
            .catch((err) => {
                apErrorFunction(err, "getApData");
            });
        let RevListName = "Review Log"
        sharepointWeb.lists
            .getByTitle(RevListName)
            .items.select(
                "*",
                "FromUser/Title",
                "FromUser/Id",
                "FromUser/EMail",
                "CcEmail/Title",
                "CcEmail/Id",
                "CcEmail/EMail"
            )
            .expand("FromUser,CcEmail")
            .top(5000)
            .orderBy("Modified", false)
            .get()
            .then(async (item) => {
                // item.forEach((x) => {
                //     const annualPlanID = x?.AnnualPlanID;
                //     if (annualPlanID && annualPlanID !== '' && Number(annualPlanID) !== 0) {
                //         const isUnique = listOfUniqueAnualIds.some(id => id === Number(annualPlanID));
                //         if (isUnique) {
                //             ReviewLogAllItems.push({
                //                 AnnualPlanID: Number(annualPlanID),
                //                 Diliverable: x?.Title,
                //                 Devoloper: x?.FromUser?.Title,
                //                 Link: "image/link",
                //                 UserName:x?.FromUser?.Title,
                //                 UserEmail:x?.FromUser?.EMail,
                //             });
                //         }
                //     }
                // });

                const listOfUser = []
                // **** apply here condiotn of response type and request  table name 
                ///////////////temprary commited
                item.map((x) => {
                    ;
                    if (listOfUniqueAnualIds.some((name) => name?.trim()?.toLowerCase() === x?.FromUser?.Title?.trim().toLowerCase() && x?.auditRequestType == "Review")) {
                        ReviewLogAllItems.push({
                            Diliverable: `${x?.Title}`,//- ${x?.ProductName} ` ,
                            DiliverableLink: x?.auditLink,
                            Devoloper: x?.FromUser?.Title,
                            UserName: x?.FromUser?.Title,
                            UserEmail: x?.FromUser?.EMail,
                            User_id: x?.FromUserId,
                            Response: x?.auditResponseType
                        });

                        listOfUser.push({
                            Devoloper: x?.FromUser?.Title,
                            UserName: x?.FromUser?.Title,
                            UserEmail: x?.FromUser?.EMail,
                            User_id: x?.FromUserId,
                        })
                    }
                })


                //understand it how its working
                const uniqueUsers = listOfUser.filter((user, index, self) =>
                    index === self.findIndex((u) => (
                        u.User_id === user.User_id
                    ))
                );



                console.log(uniqueUsers, "listOfUser");
                setAllUserList(uniqueUsers)

                setreviewLogUnsortedData(ReviewLogAllItems)
                sp.web.currentUser.get().then(user => {
                    console.log(user, "listOfUser");

                    let reviewLogSortedDataForUserOnly = []


                    ReviewLogAllItems.map((xyz) => {
                        if (xyz?.Devoloper?.trim()?.toLowerCase() == user?.Title?.trim().toLowerCase()) {

                            reviewLogSortedDataForUserOnly.push(xyz)
                        }
                    })





                    setreviewLogSortedData(reviewLogSortedDataForUserOnly);

                    setreviewLogSortedDataToBeDisplayed(reviewLogSortedDataForUserOnly);
                });


                //    setReRenderState(!ReRenderState);




                // let tempFilterKey = apFilterOptions;
                // tempFilterKey.CurrentWeek = "All";
                // tempFilterKey.CurrentMonth = "All";
                // tempFilterKey.CurrentDay = "All";
                // setApFilterOptions(tempFilterKey);

                // setApUnsortMasterData([...userOnlyDataForapAllitems]);
                // columnSortArr = ([...userOnlyDataForapAllitems]);
                // setApData([...userOnlyDataForapAllitems]);
                // columnSortMasterArr = ([...userOnlyDataForapAllitems]);
                // setApMasterData([...userOnlyDataForapAllitems]);
                // paginate(1);

                // columnSortArr = ([...userOnlyDataForapAllitems]);
                // setApData([...userOnlyDataForapAllitems]);
                // columnSortMasterArr = ([...userOnlyDataForapAllitems]);
                // setApMasterData([...userOnlyDataForapAllitems]);
                // filterKeysAfterModified(apMasterData);
                // setApFilterOptions({ ...apFilterKeys });
                // paginatewithdata(1, [...userOnlyDataForapAllitems]);




                // const testDrive = userOnlyDataForapAllitems.filter((arr) => {
                //     return arr.Year == Pb_Year;
                // });
                // let tempArrTestDrive = [...testDrive];
                // paginatewithdata(1, [...tempArrTestDrive]);
                // setApData(tempArrTestDrive)


                setApStartUpLoader(false);
            })
            .catch((err) => {
                console.log(err, "drReallocateFunction-getItem");
            });
    };
    const checkForTOD = (todType) => {
        setApResponseData({
            ID: null,
            businessArea: "",
            typeOfProject:
                apModalBoxDropDownOptions.todOptns.filter((option) => {
                    return option.key == todType;
                }).length > 0
                    ? apModalBoxDropDownOptions.todOptns.filter((option) => {
                        return option.key == todType;
                    })[0].key
                    : "",
            term: [],
            product: "",
            startDate: new Date(),
            endDate: new Date(),
            projectOrTask: "",
            Priority: "",
            ProjectVersion: "",
            year: "",
            manager: "",
            developer: [],
            status: "",
        });
        setApModalBoxVisibility({
            condition: true,
            action: "Add",
            selectedItem: [],
        });
    };
    const allitemsArrayFormatter = (item, allItems) => {


        let apDevelopersNames = [];
        let listOfDevIds = []
        let arrTerm = [];
        arrTerm.push(`${item.Term}`);
        if (item.ProjectLeadId != null) {
            item.ProjectLead.forEach((dev) => {
                listOfDevIds.push(dev.Title);
                apDevelopersNames.push({
                    name: dev.Title,
                    id: dev.Id,
                    email: dev.EMail,
                    userDetails: allPeoples.filter((people) => {
                        return people.ID == dev.Id;
                    })[0],
                });
            });
        } else {
            apDevelopersNames.push({
                name: null,
                id: null,
                email: null,
            });
        }



        var data = ({
            ID: item.ID ? item.ID : "",
            Hours: item.AllocatedHours ? item.AllocatedHours : "",
            DefaultStartDate: item.StartDate
                ? moment(item["FieldValuesAsText"].StartDate, DateListFormat).format(
                    DateListFormat
                )
                : "",
            StartDate: item.Modified
                ? moment(item["FieldValuesAsText"].Modified, DateListFormat).format(
                    DateListFormat
                )
                : "",
            DefaultEndDate: item.PlannedEndDate
                ? moment(
                    item["FieldValuesAsText"].PlannedEndDate,
                    DateListFormat
                ).format(DateListFormat)
                : "",
            EndDate: item.PlannedEndDate
                ? moment(
                    item["FieldValuesAsText"].PlannedEndDate,
                    DateListFormat
                ).format(DateListFormat)
                : "",
            Product: item.Master_x0020_ProjectId
                ? item.Master_x0020_Project.Title +
                " " +
                (item.Master_x0020_Project.ProductVersion
                    ? item.Master_x0020_Project.ProductVersion
                    : "V1")
                : "",
            TypeOfProject: item.ProjectType ? item.ProjectType : "",
            Year: item.Year ? item.Year : "",
            // Term:
            //   item.TermNew != null && item.TermNew.length > 0
            //     ? [...item.TermNew]
            //     : [],
            Term:
                item.TermNew != null && item.TermNew.length > 0
                    ? [...item.TermNew]
                    : item.Term
                        ? [...arrTerm]
                        : [],
            BusinessArea: item.BusinessArea ? item.BusinessArea : "",
            BAacronyms: item.BA_x0020_acronyms ? item.BA_x0020_acronyms : "",
            ProjectOrTask: item.Title ? item.Title : "",
            Status: item.Status ? item.Status : "",
            StatusStage: item.Status,
            DPAP: "",
            PMName:
                item.ProjectOwnerId != null
                    ? {
                        name: item.ProjectOwner.Title,
                        id: item.ProjectOwner.Id,
                        email: item.ProjectOwner.EMail,
                    }
                    : {
                        name: null,
                        id: null,
                        email: null,
                    },
            Priority: item.Priority ? item.Priority : "",
            ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
            DNames: item.ProjectLeadId != null ? [...apDevelopersNames] : [],
            listOfDevIds: listOfDevIds,
        });


        // allItems.push(data);

        return data;
    };
    const getAllOptions = () => {
        const _sortFilterKeys = (a, b) => {
            if (a.text.toLowerCase() < b.text.toLowerCase()) {
                return -1;
            }
            if (a.text.toLowerCase() > b.text.toLowerCase()) {
                return 1;
            }
            return 0;
        };

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
                            apModalBoxDrpDwnOptns.productOptns.findIndex((productOptn) => {
                                return productOptn.text == product.Title;
                            }) == -1
                        ) {
                            if (product.Title != "Not Sure") {
                                apModalBoxDrpDwnOptns.productOptns.push({
                                    key: product.Title + " " + product.ProductVersion,
                                    text: product.Title + " " + product.ProductVersion,
                                });
                            }
                            apMasterProductCollection.push({
                                productName: product.Title,
                                ProductId: product.Id,
                                ProductKey: product.Title + " " + product.ProductVersion,
                            });
                        }
                    }
                });
            })
            .then(() => {
                apModalBoxDrpDwnOptns.productOptns.sort(_sortFilterKeys);
                apModalBoxDrpDwnOptns.productOptns.unshift({
                    key: "Not Sure V1",
                    text: "Not Sure V1",
                });
            })
            .catch((err) => {
                apErrorFunction(err, "getAllOptions-Product");
            });

        //Business Area Choices
        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .fields.getByInternalNameOrTitle("BusinessArea")()
            .then((response) => {
                response["Choices"].forEach((choice) => {
                    if (choice != null) {
                        if (
                            apModalBoxDrpDwnOptns.baOptns.findIndex((baOptn) => {
                                return baOptn.key == choice;
                            }) == -1
                        ) {
                            apModalBoxDrpDwnOptns.baOptns.push({
                                key: choice,
                                text: choice,
                            });
                        }
                    }
                });
            })
            .then(() => {
                apModalBoxDrpDwnOptns.baOptns.sort(_sortFilterKeys);
            })
            .catch((err) => {
                apErrorFunction(err, "getAllOptions-Business Area");
            });

        //Priority  Choices
        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .fields.getByInternalNameOrTitle("Priority")()
            .then((response) => {
                response["Choices"].forEach((choice) => {
                    if (choice != null) {
                        if (
                            apModalBoxDrpDwnOptns.PriorityOptns.findIndex((PriorityOptn) => {
                                return PriorityOptn.key == choice;
                            }) == -1
                        ) {
                            apModalBoxDrpDwnOptns.PriorityOptns.push({
                                key: choice,
                                text: choice,
                            });
                        }
                    }
                });
            })
            .then(() => {
                apModalBoxDrpDwnOptns.PriorityOptns.sort(_sortFilterKeys);
            })
            .catch((err) => {
                apErrorFunction(err, "getAllOptions-Priority");
            });

        //Type of Deliverable Choices
        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .fields.getByInternalNameOrTitle("ProjectType")()
            .then((response) => {
                response["Choices"].forEach((choice) => {
                    let curTOD = TODacronymsCollection.filter((TOD) => {
                        return TOD.ShortName == choice;
                    });
                    let choiceText =
                        curTOD.length > 0 ? choice + " - " + curTOD[0].Name : choice;
                    if (choice != null) {
                        if (
                            apModalBoxDrpDwnOptns.todOptns.findIndex((todOptn) => {
                                return todOptn.key == choice;
                            }) == -1
                        ) {
                            apModalBoxDrpDwnOptns.todOptns.push({
                                key: choice,
                                text: choiceText,
                            });
                        }
                    }
                });
            })
            .then(() => {
                //apModalBoxDrpDwnOptns.todOptns.sort(_sortFilterKeys);
            })
            .catch((err) => {
                apErrorFunction(err, "getAllOptions-Type of Deliverable");
            });
        //Year Choices
        for (let year = moment().year(); year <= moment().year() + 10; year++) {
            apModalBoxDrpDwnOptns.yearOptns.push({
                key: year,
                text: year,
            });
        }
        //Term Choices
        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .fields.getByInternalNameOrTitle("Term")()
            .then((response) => {
                apModalBoxDrpDwnOptns.termOptns = [];
                ["1", "2", "3", "4"].forEach((choice) => {
                    if (choice != null) {
                        if (
                            apModalBoxDrpDwnOptns.termOptns.findIndex((termOptn) => {
                                return termOptn.key == choice;
                            }) == -1
                        ) {
                            apModalBoxDrpDwnOptns.termOptns.push({
                                key: choice,
                                text: choice,
                            });
                        }
                    }
                });
            })
            .then(() => {
                setApMasterProducts(apMasterProductCollection);
                setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
            })
            .catch((err) => {
                apErrorFunction(err, "getAllOptions-Term");
            });
    };
    const filterKeys = (items) => {

        items.forEach((item) => {
            if (
                apDrpDwnOptns.baOptns.findIndex((baOptn) => {
                    return baOptn.key == item.BusinessArea;
                }) == -1 &&
                item.BusinessArea
            ) {
                apDrpDwnOptns.baOptns.push({
                    key: item.BusinessArea,
                    text: item.BusinessArea,
                });
            }

            if (
                apDrpDwnOptns.todOptns.findIndex((todOptn) => {
                    return todOptn.key == item.ProjectType;
                }) == -1 &&
                item.ProjectType
            ) {
                apDrpDwnOptns.todOptns.push({
                    key: item.ProjectType,
                    text: item.ProjectType,
                });
            }
            if (
                apDrpDwnOptns.PriorityOptns.findIndex((PriorityOptn) => {
                    return PriorityOptn.key == item.Priority;
                }) == -1 &&
                item.Priority
            ) {
                apDrpDwnOptns.PriorityOptns.push({
                    key: item.Priority,
                    text: item.Priority,
                });
            }
            if (
                apDrpDwnOptns.yearOptns.findIndex((year) => {
                    return year.key == item.Year;
                }) == -1 &&
                item.Year
            ) {
                apDrpDwnOptns.yearOptns.push({
                    key: item.Year,
                    text: item.Year,
                });



            }

            if (
                apDrpDwnOptns.potOptns.findIndex((potOptn) => {
                    return potOptn.key == item.Title;
                }) == -1 &&
                item.Title
            ) {
                apDrpDwnOptns.potOptns.push({
                    key: item.Title,
                    text: item.Title,
                });
                apModalBoxDrpDwnOptns.potOptns.push({
                    key: item.Title,
                    text: item.Title,
                });
            }

            let tempmanager =
                item.ProjectOwnerId != null ? item.ProjectOwner.Title : null;
            if (
                apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
                    return managerOptn.key == tempmanager;
                }) == -1 &&
                tempmanager
            ) {
                apDrpDwnOptns.managerOptns.push({
                    key: tempmanager,
                    text: tempmanager,
                });
            }

            let tempdevelopers = [];
            if (item.ProjectLeadId != null) {
                item.ProjectLead.forEach((dev) => {
                    tempdevelopers.push(dev.Title);
                });

                tempdevelopers.forEach((tempdev) => {
                    if (
                        apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
                            return developerOptn.key == tempdev;
                        }) == -1 &&
                        tempdev
                    ) {
                        apDrpDwnOptns.developerOptns.push({
                            key: tempdev,
                            text: tempdev,
                        });
                    }
                });
            }

            // if (
            //   apDrpDwnOptns.termOptns.findIndex((termOptn) => {
            //     return termOptn.key == item.Term;
            //   }) == -1 &&
            //   item.Term
            // ) {
            //   apDrpDwnOptns.termOptns.push({
            //     key: item.Term,
            //     text: item.Term,
            //   });
            // }
        });
        ["1", "2", "3", "4"].forEach((_item) => {
            if (
                apDrpDwnOptns.termOptns.findIndex((termOptn) => {
                    return termOptn.key == _item;
                }) == -1 &&
                _item
            ) {
                apDrpDwnOptns.termOptns.push({
                    key: _item,
                    text: _item,
                });
            }
        });

        apDrpDwnOptns.yearOptns.sort((a, b) => a.key - b.key);

        const uniqueYearsSet = new Set(apDrpDwnOptns.yearOptns.map((year) => year.key));
        const uniqueYearsArray = Array.from(uniqueYearsSet).map((key) => {
            return { key: key, text: key.toString() };
        });


        const weekNames = [];
        const firstDayOfYear = moment().year(apFilterOptions.Year).startOf('year');
        const lastDayOfYear = moment().year(apFilterOptions.Year).endOf('year');
        let currentWeek = firstDayOfYear.clone().startOf('isoWeek');
        let index = 1;

        while (currentWeek.isBefore(lastDayOfYear)) {
            const weekObj = {
                key: index,
                text: currentWeek.format('W')
            };
            weekNames.push(weekObj);
            currentWeek.add(1, 'weeks');
            index++;
        }

        apDrpDwnOptns.yearOptns = uniqueYearsArray;
        apDrpDwnOptns.weekOptns = weekNames;
        let currentWeekTry = 1;

        //  setWeekDropDowns(weekNames)
        setYearDropDown(uniqueYearsArray)

        setApFilterOptions({
            ...apFilterOptions, Year: uniqueYearsArray[uniqueYearsArray.length - 1].key,
        })

        // console.log("Compilation  time", weekNames);


        sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

        setApDropDownOptions(apDrpDwnOptns);
        setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
        let bussinessOptions = apDrpDwnOptns.baOptns;
        setBussinessDropDownoptions([...bussinessOptions])
    };
    const filterKeysAfterModified = (items) => {
        items.forEach((item) => {
            if (
                apDrpDwnOptns.baOptns.findIndex((baOptn) => {
                    return baOptn.key == item.BusinessArea;
                }) == -1 &&
                item.BusinessArea
            ) {
                apDrpDwnOptns.baOptns.push({
                    key: item.BusinessArea,
                    text: item.BusinessArea,
                });
            }

            if (
                apDrpDwnOptns.todOptns.findIndex((todOptn) => {
                    return todOptn.key == item.TypeOfProject;
                }) == -1 &&
                item.TypeOfProject
            ) {
                apDrpDwnOptns.todOptns.push({
                    key: item.TypeOfProject,
                    text: item.TypeOfProject,
                });
            }
            if (
                apDrpDwnOptns.yearOptns.findIndex((year) => {
                    return year.key == item.Year;
                }) == -1 &&
                item.Year
            ) {
                apDrpDwnOptns.yearOptns.push({
                    key: item.Year,
                    text: item.Year,
                });
            }
            if (
                apDrpDwnOptns.potOptns.findIndex((potOptn) => {
                    return potOptn.key == item.ProjectOrTask;
                }) == -1 &&
                item.ProjectOrTask
            ) {
                apDrpDwnOptns.potOptns.push({
                    key: item.ProjectOrTask,
                    text: item.ProjectOrTask,
                });
                apModalBoxDrpDwnOptns.potOptns.push({
                    key: item.ProjectOrTask,
                    text: item.ProjectOrTask,
                });
            }

            let tempmanager = item.PMName != null ? item.PMName.name : null;
            if (
                apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
                    return managerOptn.key == tempmanager;
                }) == -1 &&
                tempmanager
            ) {
                apDrpDwnOptns.managerOptns.push({
                    key: tempmanager,
                    text: tempmanager,
                });
            }

            let tempdevelopers = [];
            if (item.DNames.length > 0) {
                item.DNames.forEach((dev) => {
                    tempdevelopers.push(dev.name);
                });

                tempdevelopers.forEach((tempdev) => {
                    if (
                        apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
                            return developerOptn.key == tempdev;
                        }) == -1 &&
                        tempdev != null
                    ) {
                        apDrpDwnOptns.developerOptns.push({
                            key: tempdev,
                            text: tempdev,
                        });
                    }
                });
            }

            // if (
            //   apDrpDwnOptns.termOptns.findIndex((termOptn) => {
            //     return termOptn.key == item.Term;
            //   }) == -1 &&
            //   item.Term
            // ) {
            //   apDrpDwnOptns.termOptns.push({
            //     key: item.Term,
            //     text: item.Term,
            //   });
            // }
        });
        ["1", "2", "3", "4"].forEach((_item) => {
            if (
                apDrpDwnOptns.termOptns.findIndex((termOptn) => {
                    return termOptn.key == _item;
                }) == -1 &&
                _item
            ) {
                apDrpDwnOptns.termOptns.push({
                    key: _item,
                    text: _item,
                });
            }
        });

        sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

        setApDropDownOptions(apDrpDwnOptns);
        let tempArr = apModalBoxDropDownOptions;
        tempArr.potOptns = apModalBoxDrpDwnOptns.potOptns;
        setApModalBoxDropDownOptions(tempArr);
    };
    const sortingFilterKeys = (apDrpDwnOptns, apModalBoxDrpDwnOptns) => {
        const sortFilterKeys = (a, b) => {
            if (a.text < b.text) {
                return -1;
            }
            if (a.text > b.text) {
                return 1;
            }
            return 0;
        };

        apDrpDwnOptns.baOptns.shift();
        apDrpDwnOptns.baOptns.sort(sortFilterKeys);
        apDrpDwnOptns.baOptns.unshift({ key: "All", text: "All" });

        apDrpDwnOptns.todOptns.shift();
        apDrpDwnOptns.todOptns.sort(sortFilterKeys);
        apDrpDwnOptns.todOptns.unshift({ key: "All", text: "All" });

        apDrpDwnOptns.potOptns.shift();
        apDrpDwnOptns.potOptns.sort(sortFilterKeys);
        apDrpDwnOptns.potOptns.unshift({ key: "All", text: "All" });

        apModalBoxDrpDwnOptns.potOptns.sort(sortFilterKeys);

        if (
            apDrpDwnOptns.managerOptns.some((managerOptn) => {
                return (
                    managerOptn.text.toLowerCase() ==
                    props.spcontext.pageContext.user.displayName.toLowerCase()
                );
            })
        ) {
            apDrpDwnOptns.managerOptns.shift();
            let loginUserIndex = apDrpDwnOptns.managerOptns.findIndex((user) => {
                return (
                    user.text.toLowerCase() ==
                    props.spcontext.pageContext.user.displayName.toLowerCase()
                );
            });
            let loginUserData = apDrpDwnOptns.managerOptns.splice(loginUserIndex, 1);

            apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
            apDrpDwnOptns.managerOptns.unshift(loginUserData[0]);
            apDrpDwnOptns.managerOptns = usersOrderFunction(
                apDrpDwnOptns.managerOptns
            );
            apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
        } else {
            apDrpDwnOptns.managerOptns.shift();
            apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
            apDrpDwnOptns.managerOptns = usersOrderFunction(
                apDrpDwnOptns.managerOptns
            );
            apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
        }

        if (
            apDrpDwnOptns.developerOptns.some((developerOptn) => {
                return (
                    developerOptn.text.toLowerCase() ==
                    props.spcontext.pageContext.user.displayName.toLowerCase()
                );
            })
        ) {
            apDrpDwnOptns.developerOptns.shift();
            let loginUserIndex = apDrpDwnOptns.developerOptns.findIndex((user) => {
                return (
                    user.text.toLowerCase() ==
                    props.spcontext.pageContext.user.displayName.toLowerCase()
                );
            });
            let loginUserData = apDrpDwnOptns.developerOptns.splice(
                loginUserIndex,
                1
            );
            apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
            apDrpDwnOptns.developerOptns.unshift(loginUserData[0]);
            apDrpDwnOptns.developerOptns = usersOrderFunction(
                apDrpDwnOptns.developerOptns
            );
            apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
        } else {
            apDrpDwnOptns.developerOptns.shift();
            apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
            apDrpDwnOptns.developerOptns = usersOrderFunction(
                apDrpDwnOptns.developerOptns
            );
            apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
        }

        apDrpDwnOptns.termOptns.shift();
        apDrpDwnOptns.termOptns.sort(sortFilterKeys);
        apDrpDwnOptns.termOptns.unshift({ key: "All", text: "All" });
    };
    const usersOrderFunction = (dropDown): any => {
        let nonArchived = dropDown.filter((user) => {
            return !user.text.includes("Archive");
        });
        let archived = dropDown.filter((user) => {
            return user.text.includes("Archive");
        });

        return nonArchived.concat(archived);
    };
    const listFilter = (key, option) => {
        let tempArr = [...apMasterData];
        let tempApFilterKeys = { ...apFilterOptions };
        if (key == "Week" || key == "Year" || key == "Product" || key == "BusinessArea"
        ) {
            tempApFilterKeys.CurrentDay = "All";
            tempApFilterKeys.CurrentWeek = "All";
            tempApFilterKeys.CurrentMonth = "All";

        }
        if (key == "Year" || key == "CurrentMonth") {
            tempApFilterKeys.Week = "All"
        }
        if (key == "CurrentMonth") {
            tempApFilterKeys.Week = "All";
            tempApFilterKeys.CurrentDay = "All";
            tempApFilterKeys.CurrentWeek = "All";
            tempApFilterKeys.BusinessArea = "All";
            tempApFilterKeys.Product = "All";
            tempApFilterKeys.Year = Pb_Year;
            const now = moment();
            const startOfMonth = now.startOf('month').format('DD/MM/YYYY');
            const endOfMonth = now.endOf('month').format('DD/MM/YYYY');

            const filteredList3 = tempArr.filter(item => {
                const itemStartDate = moment(item.StartDate, DateListFormat);
                if (itemStartDate.isBefore(moment(startOfMonth, DateListFormat)) || itemStartDate.isAfter(moment(endOfMonth, DateListFormat))) {
                    // Skip this item
                } else {
                    return itemStartDate.isBetween(moment(startOfMonth, DateListFormat), moment(endOfMonth, DateListFormat), null, '[]');
                }
            });
            tempArr = [...filteredList3];


        }

        if (key == "CurrentWeek") {
            tempApFilterKeys.Year = Pb_Year;


            tempApFilterKeys.CurrentDay = "All";
            tempApFilterKeys.CurrentMonth = "All";
            tempApFilterKeys.BusinessArea = "All";
            tempApFilterKeys.Product = "All";

            const today = moment();
            const currentWeek = today.isoWeek();
            const startDate = today.startOf('isoWeek').format('DD/MM/YYYY');
            const endDate = today.endOf('isoWeek').format('DD/MM/YYYY');

            tempApFilterKeys.Week = `All`;
            const filteredList4 = tempArr.filter(item => {
                const itemStartDate = moment(item.StartDate, DateListFormat);

                if (itemStartDate.isBefore(moment(startDate, DateListFormat)) || itemStartDate.isAfter(moment(endDate, DateListFormat))) {
                    // Skip this item
                } else {
                    return itemStartDate.isBetween(moment(startDate, DateListFormat), moment(endDate, DateListFormat), null, '[]');
                }

            });
            tempArr = [...filteredList4];
        }


        if (key == "CurrentDay") {
            tempApFilterKeys.Year = Pb_Year;
            tempApFilterKeys.CurrentWeek = "All";
            tempApFilterKeys.CurrentMonth = "All";
            tempApFilterKeys.BusinessArea = "All";
            tempApFilterKeys.Product = "All";

            const today = moment();
            const currentWeek = today.isoWeek();
            tempApFilterKeys.Week = `All`;


            const currentDay = moment().format('DD/MM/YYYY');

            const filteredList5 = tempArr.filter(item => {
                const itemStartDate = moment(item.StartDate, DateListFormat);

                if (itemStartDate.isBefore(moment(currentDay, DateListFormat)) || itemStartDate.isAfter(moment(currentDay, DateListFormat))) {
                } else {
                    return itemStartDate.isBetween(moment(currentDay, DateListFormat), currentDay, null, '[]');
                }

                if (!itemStartDate.isSame(moment(currentDay, DateListFormat), 'day')) {
                    // Skip this item
                } else {
                    return true;
                }

            });
            tempArr = [...filteredList5];
        }
        // tempApFilterKeys['Product'] = 'All';
        tempApFilterKeys[`${key}`] = option;



        if (tempApFilterKeys.ProjectOrTaskSearch) {
            tempArr = tempArr.filter((arr) => {
                return arr.ProjectOrTask.toLowerCase().includes(
                    tempApFilterKeys.ProjectOrTaskSearch.toLowerCase()
                );
            });
        }
        if (tempApFilterKeys.BusinessArea != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.BusinessArea == tempApFilterKeys.BusinessArea;
            });
        }
        if (tempApFilterKeys.TypeOfProject != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.TypeOfProject == tempApFilterKeys.TypeOfProject;
            });
        }
        if (tempApFilterKeys.ProjectOrTask != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.ProjectOrTask == tempApFilterKeys.ProjectOrTask;
            });
        }
        if (tempApFilterKeys.PM != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.PMName.name == tempApFilterKeys.PM;
            });
        }
        if (tempApFilterKeys.Week != "All") {
            if (key != "CurrentMonth" && key != "CurrentWeek" && key != "CurrentDay") {
                const dateFirst = moment().year(tempApFilterKeys.Year).week((Number(tempApFilterKeys.Week))).startOf('week');
                const dateLast = moment().year(tempApFilterKeys.Year).week(Number(tempApFilterKeys.Week)).endOf('week');
                const startingDateOfWeek = dateFirst.format(DateListFormat);
                const endingDateOfWeek = dateLast.format(DateListFormat);
                //   const endingDateOfYear = moment(startingDateOfWeek, DateListFormat).endOf('year');

                const filteredList = tempArr.filter(item => {

                    const itemStartDate = moment(item.StartDate, DateListFormat);
                    if (itemStartDate.isBefore(moment(startingDateOfWeek, DateListFormat)) || itemStartDate.isAfter(moment(endingDateOfWeek, DateListFormat))) {
                    } else {
                        return itemStartDate.isBetween(moment(startingDateOfWeek, DateListFormat), moment(endingDateOfWeek, DateListFormat), null, '[]');
                    }



                });
                tempArr = [...filteredList];

                //apply here condition to fetch data only between selected week start and ending date---------
                // tempArr = tempArr.filter((arr) => {
                //     return arr.PMName.name == tempApFilterKeys.PM;
                // });
            }

        } else {


            if (key != "CurrentMonth" && key != "CurrentWeek" && key != "CurrentDay") {

                const tempData = tempArr.filter((arr) => {
                    return arr.Year == tempApFilterKeys.Year;
                });
                tempArr = [...tempData];

            }

        }

        if (tempApFilterKeys.D != "All") {
            let devArr = [];
            tempArr.forEach((arr) => {
                if (arr.DNames.length != 0) {

                    if (arr.DNames.some((DName) => DName.name == option)) {
                        devArr.push(arr);

                    }
                }

                // if (tempApFilterKeys != null && pageSwitch == "User") {
                //     if (data.DNames.length != 0) {
                //         if (data.DNames.some((DName) => DName.name == tempApFilterKeys)) {
                //             allItems.push(data);
                //             console.log("filtered Data");

                //         }
                //     }





            });
            tempArr = [...devArr];
        }
        if (tempApFilterKeys.Term != "All") {
            let termArr = [];
            tempArr.forEach((arr) => {
                if (arr.Term.length != 0) {
                    if (arr.Term.some((term) => term == tempApFilterKeys.Term)) {
                        termArr.push(arr);
                    }
                }
            });
            tempArr = [...termArr];
        }
        if (tempApFilterKeys.Product != "All") {

            let termArr = [];
            tempArr.forEach((arr) => {

                if (arr.Product == tempApFilterKeys.Product) {
                    termArr.push(arr);
                }

            });
            tempArr = [...termArr];
        }


        filterKeysAfterModified([...tempArr]);
        paginatewithdata(1, [...tempArr]);
        setApFilterOptions(tempApFilterKeys);
        columnSortArr = [...tempArr];
        setApData([...tempArr]);


        if (pageSwitch == "User") {

            setreviewLogSortedDataToBeDisplayed(reviewLogSortedData)
            //add here if no  recird availbe remove anomilies also
        } else {
            let allReviewLogs = reviewLogUnsortedData;
            var listOfUniqueAnualIds2: number[] = [];

            // finding user in each product 
            tempArr.map((zz) => {



                zz?.listOfDevIds.map((x) => {
                    if (
                        listOfUniqueAnualIds2.findIndex((baOptn) => {
                            return baOptn == x;
                        }) == -1
                    ) {
                        listOfUniqueAnualIds2.push(x)
                    }
                })

            })



            let ReviewLogAllItems = []



            allReviewLogs.map((x) => {
                if (listOfUniqueAnualIds2.some((name) => name == x.Devoloper)) {
                    ReviewLogAllItems.push({
                        DiliverableLink: x.DiliverableLink,
                        Diliverable: x.Diliverable,
                        Devoloper: x.Devoloper,
                        UserName: x.UserName,
                        UserEmail: x.UserEmail,
                        User_id: x.User_id,
                        Response: x.Response,
                    });
                }
            })



            setreviewLogSortedDataToBeDisplayed(ReviewLogAllItems);


            // assign this newly generated data to display screen
        }
        return [...tempArr]


    };


    const listFilterAfterUpdated = (masterData: any) => {
        let tempArr = [...masterData];
        let tempApFilterKeys = { ...apFilterOptions };

        if (tempApFilterKeys.ProjectOrTaskSearch) {
            tempArr = tempArr.filter((arr) => {
                return arr.ProjectOrTask.toLowerCase().includes(
                    tempApFilterKeys.ProjectOrTaskSearch.toLowerCase()
                );
            });
        }
        if (tempApFilterKeys.BusinessArea != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.BusinessArea == tempApFilterKeys.BusinessArea;
            });
        }
        if (tempApFilterKeys.TypeOfProject != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.TypeOfProject == tempApFilterKeys.TypeOfProject;
            });
        }
        if (tempApFilterKeys.ProjectOrTask != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.ProjectOrTask == tempApFilterKeys.ProjectOrTask;
            });
        }
        if (tempApFilterKeys.PM != "All") {
            tempArr = tempArr.filter((arr) => {
                return arr.PMName.name == tempApFilterKeys.PM;
            });
        }
        // if (tempApFilterKeys.Year != "All") {
        tempArr = tempArr.filter((arr) => {
            return arr.Year == tempApFilterKeys.Year;
        });
        // }
        if (tempApFilterKeys.D != "All") {
            let devArr = [];
            tempArr.forEach((arr) => {
                if (arr.DNames.length != 0) {
                    if (arr.DNames.some((DName) => DName.name == tempApFilterKeys.D)) {
                        devArr.push(arr);

                    }
                }
            });
            tempArr = [...devArr];
        }
        if (tempApFilterKeys.Term != "All") {
            let termArr = [];
            tempArr.forEach((arr) => {

                if (arr.Product.length != 0) {
                    if (arr.Product.some((Pr) => Pr == tempApFilterKeys.Product)) {
                        termArr.push(arr);
                    }
                }
            });
            tempArr = [...termArr];
        }

        return tempArr;
    };
    const dateFormater = (date: Date): string => {
        return date ? moment(date).format("DD/MM/YYYY") : "";
    };
    const GetUserDetails = (filterText) => {
        var result = allPeoples.filter(
            (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
        );

        return result.filter((item) =>
            doesTextStartWith(item.text as string, filterText)
        );
    };
    const doesTextStartWith = (text: string, filterText: string) => {
        return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
    };
    const onChangeHandler = (key: string, value: any) => {
        console.log(key, "key", value, "value");


        let tempArr = allUserDataForapAllitems

        // setAllUserDataForapAllitems([...tempArrCustom]);

        // let tempArr = [...tempArrCustom];
        // let tempApFilterKeys = { ...apFilterOptions };
        // sp.web.currentUser.get().then(user => {

        // setcurrentUser(user)
        // console.log(user, "Current login user")
        let devArr = [];
        tempArr.forEach((arr) => {
            if (arr.DNames.length != 0) {
                if (arr.DNames.some((DName) => DName.id == value)) {
                    devArr.push(arr);
                }
            }
        });

        //login user specifc data 
        setUserOnlyDataForapAllitems([...devArr])

        setApUnsortMasterData([...devArr]);
        columnSortArr = [...devArr];
        setApData([...devArr]);
        columnSortMasterArr = [...devArr];
        setApMasterData([...devArr]);
        paginate(1);
        columnSortArr = [...devArr];
        setApData([...devArr]);
        columnSortMasterArr = [...devArr];
        setApMasterData([...devArr]);
        setMasterApColumn(apColumns);
        filterKeysAfterModified(apMasterData);
        paginatewithdata(1, [...devArr]);
        setApFilterOptions({ ...apFilterKeys });

        const currentYearData = devArr.filter((arr) => {
            return arr.Year == Pb_Year;
        });
        let tempArrTestDrive = [...currentYearData];
        paginatewithdata(1, [...tempArrTestDrive]);
        setApData(tempArrTestDrive)
        let ReviewLogAllItems = reviewLogUnsortedData
        let reviewLogSortedDataForUserOnly = []
        ReviewLogAllItems.map((xyz) => {
            if (xyz?.User_id == value) {
                reviewLogSortedDataForUserOnly.push(xyz)
            }
        })
        setreviewLogSortedData(reviewLogSortedDataForUserOnly);
        setreviewLogSortedDataToBeDisplayed(reviewLogSortedDataForUserOnly);

    };
    const apAddItem = () => {
        let product = [];
        let devIds = [];
        if (apResponseData.developer.length > 0) {
            apResponseData.developer.forEach((dev) => {
                devIds.push(dev.ID);
            });
        }

        if (apResponseData.product != null) {
            product = apMasterProducts.filter((prod) => {
                return prod.ProductKey == apResponseData.product;
            });
        }

        const requestdata = {
            Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
            Status: "Scheduled",
            Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
            ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
            ProjectLeadId:
                apResponseData.developer.length > 0
                    ? { results: [...devIds] }
                    : { results: [] },
            Year: apResponseData.year ? apResponseData.year : null,
            TermNew:
                apResponseData.term.length > 0
                    ? { results: [...apResponseData.term] }
                    : { results: [] },
            BusinessArea: apResponseData.businessArea
                ? apResponseData.businessArea
                : null,
            Priority: apResponseData.Priority ? apResponseData.Priority : null,
            ProjectVersion: apResponseData.ProjectVersion
                ? apResponseData.ProjectVersion
                : "",
            BA_x0020_acronyms: apResponseData.businessArea
                ? BAacronymsCollection.filter((BAacronym) => {
                    return BAacronym.Name == apResponseData.businessArea;
                })[0].ShortName
                : null,
            ProjectType: apResponseData.typeOfProject
                ? apResponseData.typeOfProject
                : null,
            StartDate: apResponseData.startDate
                ? moment(apResponseData.startDate, DateListFormat).format("YYYY-MM-DD")
                : null,
            PlannedEndDate: apResponseData.endDate
                ? moment(apResponseData.endDate, DateListFormat).format("YYYY-MM-DD")
                : null,
        };

        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .items.add(requestdata)
            .then((e) => {
                sharepointWeb.lists
                    .getByTitle(ListNameURL)
                    .items.getById(e.data.Id)
                    .select(
                        "*",
                        "ProjectOwner/Title",
                        "ProjectOwner/Id",
                        "ProjectOwner/EMail",
                        "ProjectLead/Title",
                        "ProjectLead/Id",
                        "ProjectLead/EMail",
                        "Master_x0020_Project/Title",
                        "Master_x0020_Project/Id",
                        "Master_x0020_Project/ProductVersion",
                        "FieldValuesAsText/StartDate",
                        "FieldValuesAsText/PlannedEndDate"
                    )
                    .expand(
                        "ProjectOwner",
                        "ProjectLead",
                        "Master_x0020_Project",
                        "FieldValuesAsText"
                    )
                    .get()
                    .then((item) => {
                        let tempMasterArr = [...apMasterData];
                        let newItemAddedtoArr = [];
                        let arrAfterAddApData = allitemsArrayFormatter(
                            item,
                            newItemAddedtoArr
                        );

                        Array.prototype.push.apply(arrAfterAddApData, tempMasterArr);

                        filterKeysAfterModified(arrAfterAddApData);
                        let lastIndex: number = 1 * totalPageItems;
                        let firstIndex: number = lastIndex - totalPageItems;
                        // let paginatedItems = arrAfterAddApData.slice(firstIndex, lastIndex);

                        setApModalBoxVisibility({
                            condition: false,
                            action: "",
                            selectedItem: [],
                        });

                        // setApUnsortMasterData([...arrAfterAddApData]);
                        // columnSortArr = arrAfterAddApData;
                        // setApData(arrAfterAddApData);
                        // columnSortMasterArr = arrAfterAddApData;
                        // setApMasterData([...arrAfterAddApData]);
                        // setdisplayData([...paginatedItems]);
                        // setApCurrentPage(1);
                        // setApShowMessage(apErrorStatus);
                        // setApResponseData({ ...responseData });
                        // setApModelBoxDrpDwnToTxtBox(false);
                        // setApOnSubmitLoader(false);
                        // AddSuccessPopup();
                    })
                    .catch((err) => {
                        apErrorFunction(err, "apAddItem-getItem");
                    });
            })
            .catch((err) => {
                apErrorFunction(err, "apAddItem-updateItem");
            });
    };
    const apUpdateItem = (id: number) => {
        let product = [];
        let devIds = [];
        if (apResponseData.developer.length > 0) {
            apResponseData.developer.forEach((dev) => {
                devIds.push(dev.ID);
            });
        }

        if (apResponseData.product != null) {
            product = apMasterProducts.filter((prod) => {
                return prod.ProductKey == apResponseData.product;
            });
        }

        const requestdata = {
            Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
            Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
            ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
            ProjectLeadId:
                apResponseData.developer.length > 0
                    ? { results: [...devIds] }
                    : { results: [] },
            Year: apResponseData.year ? apResponseData.year : null,
            TermNew:
                apResponseData.term.length > 0
                    ? { results: [...apResponseData.term] }
                    : { results: [] },
            BusinessArea: apResponseData.businessArea
                ? apResponseData.businessArea
                : null,
            BA_x0020_acronyms: apResponseData.businessArea
                ? BAacronymsCollection.filter((BAacronym) => {
                    return BAacronym.Name == apResponseData.businessArea;
                })[0].ShortName
                : null,
            ProjectType: apResponseData.typeOfProject
                ? apResponseData.typeOfProject
                : null,
            Priority: apResponseData.Priority ? apResponseData.Priority : null,
            ProjectVersion: apResponseData.ProjectVersion
                ? apResponseData.ProjectVersion
                : "",
            StartDate: apResponseData.startDate
                ? moment(apResponseData.startDate, DateListFormat).format("YYYY-MM-DD")
                : null,
            PlannedEndDate: apResponseData.endDate
                ? moment(apResponseData.endDate, DateListFormat).format("YYYY-MM-DD")
                : null,
            Status:
                apResponseData.status == "Completed"
                    ? "Completed"
                    : apResponseData.status == "On hold"
                        ? "On hold"
                        : apResponseData.status,
        };

        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .items.getById(id)
            .update(requestdata)
            .then(() => {
                sharepointWeb.lists
                    .getByTitle(ListNameURL)
                    .items.getById(id)
                    .select(
                        "*",
                        "ProjectOwner/Title",
                        "ProjectOwner/Id",
                        "ProjectOwner/EMail",
                        "ProjectLead/Title",
                        "ProjectLead/Id",
                        "ProjectLead/EMail",
                        "Master_x0020_Project/Title",
                        "Master_x0020_Project/Id",
                        "Master_x0020_Project/ProductVersion",
                        "FieldValuesAsText/StartDate",
                        "FieldValuesAsText/PlannedEndDate"
                    )
                    .expand(
                        "ProjectOwner",
                        "ProjectLead",
                        "Master_x0020_Project",
                        "FieldValuesAsText"
                    )
                    .get()
                    .then((item) => {
                        let tempMasterArr = [...apMasterData];
                        let updatedItemtoArr = [];
                        let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
                        tempMasterArr.splice(targetIndex, 1);
                        let arrAfterUpdateApData = allitemsArrayFormatter(
                            item,
                            updatedItemtoArr
                        );
                        Array.prototype.push.apply(arrAfterUpdateApData, tempMasterArr);

                        let ArrAfterUpdated = listFilterAfterUpdated(arrAfterUpdateApData);

                        filterKeysAfterModified(ArrAfterUpdated);

                        let lastIndex: number = 1 * totalPageItems;
                        let firstIndex: number = lastIndex - totalPageItems;
                        let paginatedItems = ArrAfterUpdated.slice(firstIndex, lastIndex);

                        setApModalBoxVisibility({
                            condition: false,
                            action: "",
                            selectedItem: [],
                        });

                        // setApUnsortMasterData(arrAfterUpdateApData);
                        // columnSortMasterArr = arrAfterUpdateApData;
                        // setApMasterData([...arrAfterUpdateApData]);
                        columnSortArr = ArrAfterUpdated;
                        setApData(ArrAfterUpdated);

                        setdisplayData([...paginatedItems]);
                        setApCurrentPage(1);
                        setApShowMessage(apErrorStatus);
                        setApResponseData({ ...responseData });
                        setApOnSubmitLoader(false);
                        setApSubmitConfirmLoader(false);
                        setSubmitConfirmationPopup(false);
                        UpdateSuccessPopup();
                    })
                    .catch((err) => {
                        apErrorFunction(err, "apUpdateItem-updateItem");
                    });
            })
            .catch((err) => {
                apErrorFunction(err, "apUpdateItem-updateItem");
            });
    };
    const apDeleteItem = (id: number) => {
        sharepointWeb.lists
            .getByTitle(ListNameURL)
            .items.getById(id)
            .delete()
            .then(() => {
                let tempMasterArr = [...apMasterData];
                let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
                tempMasterArr.splice(targetIndex, 1);

                let temp_ap_arr = [...apData];
                let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
                temp_ap_arr.splice(targetIndexapdata, 1);

                filterKeysAfterModified(temp_ap_arr);

                setApUnsortMasterData(tempMasterArr);
                columnSortMasterArr = tempMasterArr;
                setApMasterData(tempMasterArr);
                columnSortArr = temp_ap_arr;
                setApData(temp_ap_arr);
                paginatewithdata(apcurrentPage, temp_ap_arr);
                setApOnDeleteLoader(false);
                setApDeletePopup({ condition: false, targetId: 0 });
                DeleteSuccessPopup();
            })
            .catch((err) => {
                apErrorFunction(err, "apDeleteItem");
            });
    };
    const apValidationFunction = () => {
        let isError = false;

        let errorStatus = {
            businessAreaError: "",
            projectOrTaskError: "",
            productError: "",
        };
        if (!apResponseData.product) {
            isError = true;
            errorStatus.productError = "Please select product or solution";
        }

        if (!apResponseData.businessArea) {
            isError = true;
            errorStatus.businessAreaError = "Please select business area";
        }
        if (!apResponseData.projectOrTask) {
            isError = true;
            errorStatus.projectOrTaskError = "Please select name of the deliverable";
        }

        if (!isError) {
            if (apModalBoxVisibility.action == "Add") {
                setApOnSubmitLoader(true);
                apAddItem();
            } else if (apModalBoxVisibility.action == "Update") {
                let filteredArr = apMasterData.filter((data) => {
                    return data.ID == apResponseData.ID;
                })[0];
                if (
                    apResponseData.status == "Completed" &&
                    filteredArr.Status != "Completed"
                ) {
                    setSubmitConfirmationPopup(true);
                } else {
                    setApOnSubmitLoader(true);
                    apUpdateItem(apResponseData.ID);
                }
            }
        } else {
            setApShowMessage(errorStatus);
        }
    };
    const paginate = (pagenumber) => {
        let lastIndex: number = pagenumber * totalPageItems;
        let firstIndex: number = lastIndex - totalPageItems;
        let paginatedItems = apData;
        currentpage = pagenumber;

        //setdisplayData(paginatedItems);
        setApCurrentPage(pagenumber);
    };
    const paginatewithdata = (pagenumber, data) => {
        let lastIndex: number = pagenumber * totalPageItems;
        let firstIndex: number = lastIndex - totalPageItems;
        let paginatedItems = data;
        ;

        currentpage = pagenumber;



        if (paginatedItems.length > 0) {
            setdisplayData([...paginatedItems]);
            setApCurrentPage(pagenumber);
        } else {

            setdisplayData([]);
            //  paginate(pagenumber - 1);
        }
        setreRenderApplicationData(!reRenderApplicationData)
    };

    const _onColumnClick = (
        ev: React.MouseEvent<HTMLElement>,
        column: IColumn
    ): void => {
        const tempapColumns = apColumns;
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

        const newData = _copyAndSort(
            columnSortArr,
            currColumn.fieldName!,
            currColumn.isSortedDescending
        );
        const newMasterData = _copyAndSort(
            columnSortMasterArr,
            currColumn.fieldName!,
            currColumn.isSortedDescending
        );
        setApData([...newData]);
        setApMasterData([...newMasterData]);
        paginatewithdata(1, newData);
    };
    function _copyAndSort<T>(
        items: T[],
        columnKey: string,
        isSortedDescending?: boolean
    ): T[] {
        let key = columnKey as keyof T;
        if (key == "PM") {
            const ascSortFunction = (a, b) => {
                if (a.PMName["name"] < b.PMName["name"]) {
                    return -1;
                }
                if (a.PMName["name"] > b.PMName["name"]) {
                    return 1;
                }
                return 0;
            };
            const decSortFunction = (b, a) => {
                if (a.PMName["name"] < b.PMName["name"]) {
                    return -1;
                }
                if (a.PMName["name"] > b.PMName["name"]) {
                    return 1;
                }
                return 0;
            };

            return items.sort(isSortedDescending ? ascSortFunction : decSortFunction);
        } else {
            return items
                .slice(0)
                .sort((a: T, b: T) =>
                    (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
                );
        }
    }
    const AddSuccessPopup = () => (
        alertify.set("notifier", "position", "top-right"),
        alertify.success("Annual plan is successfully submitted !!!")
    );
    const UpdateSuccessPopup = () => (
        alertify.set("notifier", "position", "top-right"),
        alertify.success("Annual plan is successfully updated !!!")
    );
    const DeleteSuccessPopup = () => (
        alertify.set("notifier", "position", "top-right"),
        alertify.success("Annual plan is successfully deleted !!!")
    );
    const ErrorPopup = () => (
        alertify.set("notifier", "position", "top-right"),
        alertify.error("Something when error, please contact system admin.")
    );
    const apErrorFunction = (error: any, functionName: string) => {
        let response = {
            ComponentName: "Annual plan",
            FunctionName: functionName,
            ErrorMessage: JSON.stringify(error["message"]),
            Title: loggeduseremail,
        };

        Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
            () => {
                setApStartUpLoader(false);
                setApOnSubmitLoader(false);
                setApOnDeleteLoader(false);
                ErrorPopup();
            }
        );
    };

    useEffect(() => {
        getAllOptions();
        getApData();
    }, [apReRender]);


    useEffect(() => {


    }, [reviewLogSortedDataToBeDisplayed])
    useEffect(() => {
        const handleResize = () => {
            if (window.innerWidth <= 800) {
                setstackItemWidth({
                    leftSec: "100%",
                    rightSection: "100%",
                    mobileView: true
                })
                setStackProps({ horizontal: false });
            } else {
                setstackItemWidth({
                    leftSec: "60%",
                    rightSection: "40%",
                    mobileView: false
                })
                setStackProps({ horizontal: true });
            }
        };

        handleResize();

        window.addEventListener('resize', handleResize);
        return () => window.removeEventListener('resize', handleResize);
    }, []);






    return (
        <div style={{ padding: "5px 15px" }}>
            {apStartUpLoader ? <CustomLoader /> : null}
            <div style={{ padding: "10px 15px" }}>
                {/* header-Section Starts*/}
                <div className={styles.apHeaderSection}>
                    <div
                        style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                        }}
                    >
                        <div className={styles.apHeader}>Dashboard</div>

                    </div>
                    <div
                        style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center",
                        }}
                    >
                        <div   className={styles.WRButtonSection}>
                            <button
                                className={
                                    pageSwitch == "User"
                                        ? styles.activeButton
                                        : styles.inactiveButton
                                }
                                onClick={() => {
                                    setHistory({
                                        condition: false,
                                        sourcePage: "",
                                        targetID: null,
                                    });

                                    const allUserData = userOnlyDataForapAllitems;
                                    setApUnsortMasterData([...allUserData]);
                                    columnSortArr = [...allUserData];
                                    columnSortMasterArr = [...allUserData];
                                    setApMasterData([...allUserData]);
                                    setMasterApColumn(apColumns);
                                    filterKeysAfterModified(apMasterData);
                                    setApFilterOptions({ ...apFilterKeys });
                                    paginatewithdata(1, [...allUserData]);
                                    const currentYearData = allUserData.filter((arr) => {
                                        return arr.Year == Pb_Year;
                                    });
                                    let tempArrTestDrive = [...currentYearData];
                                    paginatewithdata(1, [...tempArrTestDrive]);
                                    setApData(tempArrTestDrive)
                                    const allUserDRA = reviewLogSortedData;
                                    setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                    setPageSwitch("User");
                                }}


                            >
                                User
                            </button>
                            <button
                                className={
                                    pageSwitch == "Product"
                                        ? styles.activeButton
                                        : styles.inactiveButton
                                }
                                onClick={() => {
                                    setHistory({
                                        condition: false,
                                        sourcePage: "",
                                        targetID: null,
                                    });

                                    const allUserData = allUserDataForapAllitems;
                                    setApUnsortMasterData([...allUserData]);
                                    columnSortArr = [...allUserData];
                                    columnSortMasterArr = [...allUserData];
                                    setApMasterData([...allUserData]);
                                    setMasterApColumn(apColumns);
                                    filterKeysAfterModified(apMasterData);
                                    setApFilterOptions({ ...apFilterKeys });

                                    const currentYearData = allUserData.filter((arr) => {
                                        return arr.Year == Pb_Year;
                                    });
                                    let tempArrTestDrive = [...currentYearData];
                                    paginatewithdata(1, [...tempArrTestDrive]);
                                    setApData(tempArrTestDrive)
                                    const allUserDRA = reviewLogUnsortedData;
                                    setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                    setPageSwitch("Product");


                                    //  setreviewLogSortedData(reviewLogUnsortedData)
                                }}
                            >
                                Product
                            </button>
                            <button
                                className={
                                    pageSwitch == "BussinessArea"
                                        ? styles.activeButton
                                        : styles.inactiveButton
                                }
                                onClick={() => {
                                    setHistory({
                                        condition: false,
                                        sourcePage: "",
                                        targetID: null,
                                    });
                                    const allUserData = allUserDataForapAllitems;

                                    setApUnsortMasterData([...allUserData]);
                                    columnSortArr = [...allUserData];

                                    columnSortMasterArr = [...allUserData];
                                    setApMasterData([...allUserData]);
                                    setMasterApColumn(apColumns);
                                    filterKeysAfterModified(apMasterData);
                                    setApFilterOptions({ ...apFilterKeys });

                                    const currentYearData = allUserData.filter((arr) => {
                                        return arr.Year == Pb_Year;
                                    });
                                    let tempArrTestDrive = [...currentYearData];
                                    paginatewithdata(1, [...tempArrTestDrive]);
                                    setApData(tempArrTestDrive)
                                    const allUserDRA = reviewLogUnsortedData;

                                    setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                    setPageSwitch("BussinessArea");

                                }}
                            >
                                Business Area
                            </button>
                            <button
                                className={
                                    pageSwitch == "organisation"
                                        ? styles.activeButton
                                        : styles.inactiveButton
                                }
                                onClick={() => {
                                    setHistory({
                                        condition: false,
                                        sourcePage: "",
                                        targetID: null,
                                    });

                                    /////////////////////remapping data of getapData for all user
                                    const allUserData = allUserDataForapAllitems;
                                    // setUserOnlyDataForapAllitems([...allUserData])

                                    setApUnsortMasterData([...allUserData]);
                                    columnSortArr = [...allUserData];
                                    columnSortMasterArr = [...allUserData];
                                    setApMasterData([...allUserData]);
                                    setMasterApColumn(apColumns);
                                    filterKeysAfterModified(apMasterData);
                                    setApFilterOptions({ ...apFilterKeys });

                                    const currentYearData = allUserData.filter((arr) => {
                                        return arr.Year == Pb_Year;
                                    });
                                    let tempArrTestDrive = [...currentYearData];
                                    paginatewithdata(1, [...tempArrTestDrive]);
                                    setApData(tempArrTestDrive)
                                    const allUserDRA = reviewLogUnsortedData;
                                    setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                    setPageSwitch("organisation");



                                }}
                            >
                                Organisation
                            </button>
                        </div>

                        <div style={{ display: "flex", justifyContent: "space-between" }}>
                            <div style={{ display: "flex", alignItems: "center" }}>

                              {Object.keys(currentUser).length > 0 &&  <NormalPeoplePicker
                                    className={apModalBoxPP}
                                    onResolveSuggestions={GetUserDetails}
                                    itemLimit={1}
                                    defaultSelectedItems={[{
                                        // ID: 372,
                                        imageUrl: `/_layouts/15/userphoto.aspx?size=S&accountname=${currentUser["Email"]}`,
                                        // isValid: true,
                                        key: 1,
                                        secondaryText: currentUser["Email"],
                                        text: currentUser["Title"]
                                    }]
                                        // return people.secondaryText.toLowerCase() == currentUser.Email.toLowerCase(); 
                                    }
                                    styles={{
                                        root: {
                                            width: 300,
                                            margin: "10px 20px",
                                            selectors: {
                                                ".ms-BasePicker-text": {
                                                    height: 36,
                                                    padding: "3px 10px",
                                                    border: "1px solid black",
                                                    borderRadius: 4,
                                                },
                                            },
                                            ".ms-Persona-primaryText": { fontWeight: 600 },
                                        },
                                    }}
                                    onChange={(selectedUser) => {
                                        selectedUser.length != 0
                                            ? onChangeHandler("manager", selectedUser[0]["ID"])
                                            : onChangeHandler("manager", "");
                                    }}
                                />}
{/* 
                              <Label style={{ marginRight: "10px" }}>
                                    {currentUser["Title"]} - {currentUser["Email"]}
                                </Label> */}
                                 {/*  <Persona
                                    size={PersonaSize.size24}
                                    presence={PersonaPresence.none}
                                    imageUrl={
                                        "/_layouts/15/userphoto.aspx?size=S&username=" +
                                        `${currentUser["Email"]}`
                                    }
                                /> */}
                            </div>
                        </div>
                    </div>
                </div>
                {/* header-Section Ends*/}
                {/* body-Section Starts */}
                <div>
                    <>  {
                        pageSwitch == "User" ? (
                            <>
                                <div className={styles.apHeaderSection}>
                                    <div
                                        style={{
                                            display: "flex",
                                            justifyContent: "space-between",
                                            alignItems: "center",
                                        }}
                                    >
                                        <div style={{
                                            display: "flex",
                                            alignItems: "center",
                                            justifyContent: "left",
                                            paddingTop: 16,
                                        }}>

                                            <div>
                                                <Label styles={DBfilterShortLabelStyles}>Week</Label>
                                                <Dropdown
                                                    selectedKey={apFilterOptions.Week}
                                                    multiSelect={false}
                                                    placeholder="Select Week"
                                                    options={weekDropDowns}
                                                    styles={
                                                        apActiveShortDropdownStyles
                                                    }
                                                    onChange={(e, option: any) => {

                                                        listFilter("Week", option["key"]);

                                                    }}
                                                />
                                            </div>

                                            <div>
                                                <Label styles={apShortLabelStyles}>Year</Label>
                                                <Dropdown
                                                    selectedKey={apFilterOptions.Year}
                                                    multiSelect={false}
                                                    placeholder="Select year"
                                                    options={yearDropDown}
                                                    styles={apActiveShortDropdownStyles
                                                    }
                                                    onChange={(e, option: any) => {

                                                        listFilter("Year", option["key"]);


                                                    }}
                                                />
                                            </div>
                                            <div>
                                                <div>
                                                    <Icon
                                                        iconName="Refresh"
                                                        title="Click to reset"
                                                        className={apIconStyleClass.refresh}
                                                        onClick={() => {


                                                            let tempFilterKey = apFilterOptions;
                                                            tempFilterKey.CurrentWeek = "All";
                                                            tempFilterKey.CurrentMonth = "All";
                                                            tempFilterKey.CurrentDay = "All";
                                                            setApFilterOptions(tempFilterKey);

                                                            setApUnsortMasterData([...userOnlyDataForapAllitems]);
                                                            columnSortArr = ([...userOnlyDataForapAllitems]);
                                                            setApData([...userOnlyDataForapAllitems]);
                                                            columnSortMasterArr = ([...userOnlyDataForapAllitems]);
                                                            setApMasterData([...userOnlyDataForapAllitems]);
                                                            paginate(1);

                                                            columnSortArr = ([...userOnlyDataForapAllitems]);
                                                            setApData([...userOnlyDataForapAllitems]);
                                                            columnSortMasterArr = ([...userOnlyDataForapAllitems]);
                                                            setApMasterData([...userOnlyDataForapAllitems]);
                                                            filterKeysAfterModified(apMasterData);
                                                            setApFilterOptions({ ...apFilterKeys });
                                                            paginatewithdata(1, [...userOnlyDataForapAllitems]);




                                                            const testDrive = userOnlyDataForapAllitems.filter((arr) => {
                                                                return arr.Year == Pb_Year;
                                                            });
                                                            let tempArrTestDrive = [...testDrive];
                                                            paginatewithdata(1, [...tempArrTestDrive]);
                                                            setApData(tempArrTestDrive)


                                                            ///???review log data 
                                                            const oneUserDRA = reviewLogSortedData;

                                                            setreviewLogSortedDataToBeDisplayed([...oneUserDRA]);

                                                        }}
                                                    />
                                                </div>
                                            </div>
                                        </div>

                                        {/* Right Section  */}

                                        {
                                            !stackItemWidth.mobileView && <div
                                                style={{
                                                    display: "flex",
                                                    alignItems: "end",
                                                    justifyContent: "left",
                                                    paddingTop: 16,
                                                }}
                                            >
                                                <Stack horizontal styles={{ root: { display: 'flex', flexDirection: 'row', gap: 0, alignItems: 'center', marginLeft: 0 } }}>
                                                    <PrimaryButton


                                                        styles={apFilterOptions.CurrentDay == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                            listFilter("CurrentDay", "CurrentDay");
                                                        }}

                                                        text="Day" />
                                                    <PrimaryButton


                                                        styles={apFilterOptions.CurrentWeek == "All" ? buttonStylesBorderLess : activeButtonStyles} onClick={() => {
                                                            listFilter("CurrentWeek", "CurrentWeek");
                                                        }}


                                                        text="Week" />
                                                    <PrimaryButton styles={apFilterOptions.CurrentMonth == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentMonth", "CurrentMonth");
                                                    }} text="Month" />
                                                </Stack>

                                            </div>
                                        }


                                    </div>
                                    {/* left section */}


                                </div>

                            </>
                        ) : pageSwitch == "Product" ? (<>
                            <div className={styles.apHeaderSection}>
                                <div
                                    style={{
                                        display: "flex",
                                        justifyContent: "space-between",
                                        alignItems: "center",
                                    }}
                                >
                                    <div style={{
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "left",
                                        paddingTop: 16,
                                    }}>

                                        <div>
                                            <Label styles={pbLabelStyles}>Product</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Product}
                                                placeholder="Select an option"
                                                options={ProductDropDownoptions}
                                                styles={
                                                    // pbFilterOptions.Product == "All"
                                                    //     ? pbDropdownStyles
                                                    //     : 
                                                    pbActiveDropdownStyles
                                                }
                                                onChange={(e, option: any) => {
                                                    listFilter("Product", option["key"]);
                                                }}
                                            />
                                        </div>

                                        <div>
                                            <Label styles={DBfilterShortLabelStyles}>Week</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Week}
                                                multiSelect={false}
                                                placeholder="Select Week"
                                                options={weekDropDowns}
                                                styles={
                                                    apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {
                                                    listFilter("Week", option["key"]);
                                                }}
                                            />
                                        </div>

                                        <div>
                                            <Label styles={apShortLabelStyles}>Year</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Year}
                                                multiSelect={false}
                                                placeholder="Select year"
                                                options={yearDropDown}
                                                styles={apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {

                                                    listFilter("Year", option["key"]);


                                                }}
                                            />
                                        </div>


                                        <div>
                                            <div>
                                                <Icon
                                                    iconName="Refresh"
                                                    title="Click to reset"
                                                    className={apIconStyleClass.refresh}
                                                    onClick={() => {





                                                        ////////////


                                                        let tempFilterKey = apFilterOptions;
                                                        tempFilterKey.CurrentWeek = "All";
                                                        tempFilterKey.CurrentMonth = "All";
                                                        tempFilterKey.CurrentDay = "All";
                                                        setApFilterOptions(tempFilterKey);


                                                        /////////////////////remapping data of getapData for all user
                                                        const allUserData = allUserDataForapAllitems;
                                                        //   setUserOnlyDataForapAllitems([...allUserData])

                                                        setApUnsortMasterData([...allUserData]);
                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        paginate(1);

                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        setMasterApColumn(apColumns);
                                                        filterKeysAfterModified(apMasterData);
                                                        setApFilterOptions({ ...apFilterKeys });
                                                        paginatewithdata(1, [...allUserData]);




                                                        const currentYearData = allUserData.filter((arr) => {
                                                            return arr.Year == Pb_Year;
                                                        });
                                                        let tempArrTestDrive = [...currentYearData];
                                                        paginatewithdata(1, [...tempArrTestDrive]);
                                                        setApData(tempArrTestDrive)





                                                        const allUserDRA = reviewLogUnsortedData;

                                                        setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                                    }}
                                                />
                                            </div>
                                        </div>
                                    </div>

                                    {/* Right Section  */}

                                    {
                                        !stackItemWidth.mobileView && <div
                                            style={{
                                                display: "flex",
                                                alignItems: "end",
                                                justifyContent: "left",
                                                paddingTop: 16,
                                            }}
                                        >
                                            <Stack horizontal styles={{ root: { display: 'flex', flexDirection: 'row', gap: 0, alignItems: 'center', marginLeft: 0 } }}>
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentDay == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentDay", "CurrentDay");
                                                    }}

                                                    text="Day" />
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentWeek == "All" ? buttonStylesBorderLess : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentWeek", "CurrentWeek");
                                                    }}


                                                    text="Week" />
                                                <PrimaryButton styles={apFilterOptions.CurrentMonth == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                    listFilter("CurrentMonth", "CurrentMonth");
                                                }} text="Month" />
                                            </Stack>

                                        </div>
                                    }
                                </div>
                                {/* left section */}


                            </div>




                        </>
                        ) : pageSwitch == "BussinessArea" ? (

                            <>
                                <div
                                    style={{
                                        display: "flex",
                                        justifyContent: "space-between",
                                        alignItems: "center",
                                        paddingBottom: "10px",
                                    }}
                                >
                                    <div style={{
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "left",
                                        paddingTop: 16,
                                    }}>

                                        <div>
                                            <Label styles={apLabelStyles}>Business Area</Label>
                                            <Dropdown
                                                placeholder="Select a business area"
                                                options={BussinessDropDownoptions}
                                                styles={apActiveDropdownStyles
                                                }
                                                onChange={(e, option: any) => {
                                                    listFilter("BusinessArea", option["key"]);
                                                }}
                                                selectedKey={apFilterOptions.BusinessArea}
                                            />
                                        </div>
                                        <div>
                                            <Label styles={DBfilterShortLabelStyles}>Week</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Week}
                                                multiSelect={false}
                                                placeholder="Select Week"
                                                options={weekDropDowns}
                                                styles={
                                                    apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {
                                                    listFilter("Week", option["key"]);
                                                }}
                                            />
                                        </div>

                                        <div>
                                            <Label styles={apShortLabelStyles}>Year</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Year}
                                                multiSelect={false}
                                                placeholder="Select year"
                                                options={yearDropDown}
                                                styles={apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {

                                                    listFilter("Year", option["key"]);


                                                }}
                                            />
                                        </div>

                                        <div>
                                            <div>
                                                <Icon
                                                    iconName="Refresh"
                                                    title="Click to reset"
                                                    className={apIconStyleClass.refresh}
                                                    onClick={() => {

                                                        ////////////


                                                        let tempFilterKey = apFilterOptions;
                                                        tempFilterKey.CurrentWeek = "All";
                                                        tempFilterKey.CurrentMonth = "All";
                                                        tempFilterKey.CurrentDay = "All";
                                                        setApFilterOptions(tempFilterKey);


                                                        /////////////////////remapping data of getapData for all user
                                                        const allUserData = allUserDataForapAllitems;
                                                        //  setUserOnlyDataForapAllitems([...allUserData])

                                                        setApUnsortMasterData([...allUserData]);
                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        paginate(1);

                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        setMasterApColumn(apColumns);
                                                        filterKeysAfterModified(apMasterData);
                                                        setApFilterOptions({ ...apFilterKeys });
                                                        paginatewithdata(1, [...allUserData]);




                                                        const currentYearData = allUserData.filter((arr) => {
                                                            return arr.Year == Pb_Year;
                                                        });
                                                        let tempArrTestDrive = [...currentYearData];
                                                        paginatewithdata(1, [...tempArrTestDrive]);
                                                        setApData(tempArrTestDrive)





                                                        const allUserDRA = reviewLogUnsortedData;

                                                        setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                                    }}
                                                />
                                            </div>
                                        </div>
                                    </div>

                                    {/* Right Section  */}

                                    {
                                        !stackItemWidth.mobileView && <div
                                            style={{
                                                display: "flex",
                                                alignItems: "end",
                                                justifyContent: "left",
                                                paddingTop: 16,
                                            }}
                                        >
                                            <Stack horizontal styles={{ root: { display: 'flex', flexDirection: 'row', gap: 0, alignItems: 'center', marginLeft: 0 } }}>
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentDay == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentDay", "CurrentDay");
                                                    }}

                                                    text="Day" />
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentWeek == "All" ? buttonStylesBorderLess : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentWeek", "CurrentWeek");
                                                    }}


                                                    text="Week" />
                                                <PrimaryButton styles={apFilterOptions.CurrentMonth == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                    listFilter("CurrentMonth", "CurrentMonth");
                                                }} text="Month" />
                                            </Stack>

                                        </div>
                                    }
                                </div>



                            </>
                        ) : pageSwitch == "organisation" ? (
                            <>
                                <div
                                    style={{
                                        display: "flex",
                                        justifyContent: "space-between",
                                        alignItems: "center",

                                        paddingBottom: "10px",
                                    }}
                                >
                                    <div style={{
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "left",
                                        paddingTop: 16,
                                    }}>


                                        <div>
                                            <Label styles={DBfilterShortLabelStyles}>Week</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Week}
                                                multiSelect={false}
                                                placeholder="Select Week"
                                                options={weekDropDowns}
                                                styles={
                                                    apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {
                                                    listFilter("Week", option["key"]);
                                                }}
                                            />
                                        </div>

                                        <div>
                                            <Label styles={apShortLabelStyles}>Year</Label>
                                            <Dropdown
                                                selectedKey={apFilterOptions.Year}
                                                multiSelect={false}
                                                placeholder="Select year"
                                                options={yearDropDown}
                                                styles={apActiveShortDropdownStyles
                                                }
                                                onChange={(e, option: any) => {

                                                    listFilter("Year", option["key"]);


                                                }}
                                            />
                                        </div>
                                        <div>
                                            <div>
                                                <Icon
                                                    iconName="Refresh"
                                                    title="Click to reset"
                                                    className={apIconStyleClass.refresh}
                                                    onClick={() => {
                                                        let tempFilterKey = apFilterOptions;
                                                        tempFilterKey.CurrentWeek = "All";
                                                        tempFilterKey.CurrentMonth = "All";
                                                        tempFilterKey.CurrentDay = "All";
                                                        setApFilterOptions(tempFilterKey);

                                                        /////////////////////remapping data of getapData for all user
                                                        const allUserData = allUserDataForapAllitems;
                                                        //setUserOnlyDataForapAllitems([...allUserData])

                                                        setApUnsortMasterData([...allUserData]);
                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        paginate(1);

                                                        columnSortArr = [...allUserData];
                                                        setApData([...allUserData]);
                                                        columnSortMasterArr = [...allUserData];
                                                        setApMasterData([...allUserData]);
                                                        setMasterApColumn(apColumns);
                                                        filterKeysAfterModified(apMasterData);
                                                        setApFilterOptions({ ...apFilterKeys });
                                                        paginatewithdata(1, [...allUserData]);




                                                        const currentYearData = allUserData.filter((arr) => {
                                                            return arr.Year == Pb_Year;
                                                        });
                                                        let tempArrTestDrive = [...currentYearData];
                                                        paginatewithdata(1, [...tempArrTestDrive]);
                                                        setApData(tempArrTestDrive)





                                                        const allUserDRA = reviewLogUnsortedData;

                                                        setreviewLogSortedDataToBeDisplayed([...allUserDRA]);
                                                    }}
                                                />
                                            </div>
                                        </div>
                                    </div>
                                    {/* Right Section  */}

                                    {
                                        !stackItemWidth.mobileView && <div
                                            style={{
                                                display: "flex",
                                                alignItems: "end",
                                                justifyContent: "left",
                                                paddingTop: 16,
                                            }}
                                        >
                                            <Stack horizontal styles={{ root: { display: 'flex', flexDirection: 'row', gap: 0, alignItems: 'center', marginLeft: 0 } }}>
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentDay == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentDay", "CurrentDay");
                                                    }}

                                                    text="Day" />
                                                <PrimaryButton


                                                    styles={apFilterOptions.CurrentWeek == "All" ? buttonStylesBorderLess : activeButtonStyles} onClick={() => {
                                                        listFilter("CurrentWeek", "CurrentWeek");
                                                    }}


                                                    text="Week" />
                                                <PrimaryButton styles={apFilterOptions.CurrentMonth == "All" ? buttonStyles : activeButtonStyles} onClick={() => {
                                                    listFilter("CurrentMonth", "CurrentMonth");
                                                }} text="Month" />
                                            </Stack>

                                        </div>
                                    }
                                </div>

                            </>
                        ) : null
                    }

                        {displayData.length == 0 ? (
                            <div
                                style={{
                                    display: "flex",
                                    justifyContent: "center",
                                    marginTop: "15px",
                                }}
                            >
                                <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
                            </div>
                        ) : <Stack horizontal {...stackProps} tokens={{ childrenGap: 10 }}


                        >
                            <StackItem styles={{ root: { width: stackItemWidth.leftSec } }}>
                                <div  >
                                    <div style={{ marginTop: "0px", marginBottom: "8px", fontSize: "16px", fontWeight: 500 }}>
                                        Products
                                    </div>
                                    <DetailsList
                                        items={displayData}
                                        columns={masterApColumn}
                                        styles={drDetailsListStyles}
                                        setKey="set"
                                        selectionMode={SelectionMode.none}
                                        data-is-scrollable={true}
                                        onShouldVirtualize={() => {
                                            return false;
                                        }}
                                        onRenderRow={(data, defaultRender) => (
                                            <div className="red">
                                                {defaultRender({
                                                    ...data,
                                                    styles: {
                                                        root: {
                                                            background: "#fff",
                                                            selectors: {
                                                                "&:hover": {
                                                                    background: "#f3f2f1",
                                                                },
                                                            },
                                                        },
                                                    },
                                                })}
                                            </div>
                                        )}
                                    />
                                </div>
                            </StackItem>
                            <StackItem styles={{ root: { width: stackItemWidth.rightSection } }}>
                                <div >
                                    <div style={{ marginTop: "0px", marginBottom: "8px", fontSize: "16px", fontWeight: 500 }}>
                                        Document Review Anomalies
                                    </div>
                                    {
                                        reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Cancelled").length !== 0 && <>
                                            <p style={elementStyle}>
                                                Cancelled
                                            </p>
                                            <DetailsList
                                                items={reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Cancelled")
                                                }
                                                columns={DRAColumn}
                                                styles={drDetailsListStyleSmallerThreeCount}
                                                setKey="set"
                                                selectionMode={SelectionMode.none}
                                                data-is-scrollable={true}
                                                onShouldVirtualize={() => {
                                                    return false;
                                                }}
                                                onRenderRow={(data, defaultRender) => (
                                                    <div>
                                                        {defaultRender({
                                                            ...data,
                                                            styles: {
                                                                root: {
                                                                    background: "#fff",

                                                                    selectors: {
                                                                        "&:hover": {
                                                                            background: "#f3f2f1",
                                                                        },
                                                                    },
                                                                },
                                                            },
                                                        })}
                                                    </div>
                                                )}
                                            />
                                        </>
                                    }
                                    {
                                        reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Feedback").length !== 0 && <>
                                            <p style={elementStyle}>
                                                Feedback
                                            </p>
                                            <DetailsList
                                                items={reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Feedback")}
                                                columns={DRAColumn}
                                                styles={drDetailsListStyleSmallerThreeCount}
                                                setKey="set"
                                                selectionMode={SelectionMode.none}
                                                data-is-scrollable={true}
                                                onShouldVirtualize={() => {
                                                    return false;
                                                }}
                                                onRenderRow={(data, defaultRender) => (
                                                    <div>
                                                        {defaultRender({
                                                            ...data,

                                                            styles: {
                                                                root: {
                                                                    background: "#fff",

                                                                    selectors: {
                                                                        "&:hover": {
                                                                            background: "#f3f2f1",
                                                                        },
                                                                    },
                                                                },
                                                            },
                                                        })}
                                                    </div>
                                                )}
                                                onRenderMissingItem={() => <div>No data available</div>}
                                            />
                                        </>

                                    }

                                    {
                                        reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Pending").length !== 0 && <>
                                            <p style={elementStyle}>
                                                Pending
                                            </p>
                                            <DetailsList
                                                items={reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Pending")}
                                                columns={DRAColumn}
                                                styles={drDetailsListStyleSmallerThreeCount}
                                                setKey="set"
                                                selectionMode={SelectionMode.none}
                                                data-is-scrollable={true}
                                                onShouldVirtualize={() => {
                                                    return false;
                                                }}
                                                onRenderRow={(data, defaultRender) => (
                                                    <div>
                                                        {defaultRender({
                                                            ...data,
                                                            styles: {
                                                                root: {
                                                                    background: "#fff",
                                                                    selectors: {
                                                                        "&:hover": {
                                                                            background: "#f3f2f1",
                                                                        },
                                                                    },
                                                                },
                                                            },
                                                        })}
                                                    </div>
                                                )}
                                            />
                                        </>
                                    }
                                    {
                                        reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Rejected").length !== 0 && <>
                                            <p style={elementStyle}>
                                                Rejected
                                            </p>
                                            <DetailsList
                                                items={reviewLogSortedDataToBeDisplayed.filter(x => x.Response === "Rejected")}
                                                columns={DRAColumn}
                                                styles={drDetailsListStyleSmallerThreeCount}
                                                setKey="set"
                                                selectionMode={SelectionMode.none}
                                                data-is-scrollable={true}
                                                onShouldVirtualize={() => {
                                                    return false;
                                                }}
                                                onRenderRow={(data, defaultRender) => (
                                                    <div>
                                                        {defaultRender({
                                                            ...data,

                                                            styles: {
                                                                root: {
                                                                    background: "#fff",

                                                                    selectors: {
                                                                        "&:hover": {
                                                                            background: "#f3f2f1",
                                                                        },
                                                                    },
                                                                },
                                                            },
                                                        })}
                                                    </div>
                                                )}
                                            />
                                        </>
                                    }
                                </div>
                            </StackItem>
                        </Stack>}



                    </>
                </div>
                {/* body-Section Ends */}
            </div>


        </div>
    );
};

export default Dashboard;
