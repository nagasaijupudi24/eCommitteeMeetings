/* eslint-disable react/self-closing-comp */
/* eslint-disable max-lines */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";
import './CustomStyles/custom.css'
import type { IXenWpCommitteeMeetingsFormsProps } from "./IXenWpCommitteeMeetingsFormsProps";
import {
  DatePicker,
  DefaultButton,
  defaultDatePickerStrings,
  DetailsList,
  DetailsListLayoutMode,
  // DialogType,
  Dropdown,
  IColumn,
  Icon,
  IconButton,
  IDropdownOption,
  Link,
  mergeStyleSets,
  Modal,
  PrimaryButton,
  SelectionMode,
  Spinner,
  SpinnerSize,
  TextField,
} from "@fluentui/react";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/lists";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/files";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
import DateTime from "./dateComponent";

// import { escape } from '@microsoft/sp-lodash-subset';

interface CommtteeMeetingsStateProps {
  committename: any;
  committeeNameFeildValue: any;
  charimanFeildValue: any;
  convernorFeildValue: any;
  convernorData: any;
  charimanData: any;
  selectedCommitteeMembers: any;
  selectedCommitteeGuestMembers: any;
  selectedCommitteeNoteRecords: any;
  committeeMembersData: any;
  committeeGuestMembersData: any;
  committeeNoteRecordsData: any;
  committeeMemberskey: any;
  committeeGuestMemberskey: any;
  committeeNoteRecordskey: any;
  isModalOpen: any;
  isModalMOMOpen: any;
  modalMessage: any;
  meetingId: any;
  meetingDate: any;
  meetingSubject: any;
  meetingMode: any;
  meetingLink: any;
  statusNumber: any;
  auditTrail: any;

  committeeNoteRecordDropDownData: any;
  committeeNoteRecordSelectedValue: any;
  CommitteeMeetingMemberCommentsDT:any;
  pdfLink: any;
  committeNoteRecordsDropDownDataWithAllProperties: any;
  selectedMOMNoteRecord: any;
  invalidFields: any;
  MeetingStatus: any;

  isMomDraftDialogOpen: boolean;
  dialogType: any;
  draftResolutionFieldValue: any;
  btnType: any;
  departmentAlias: any;
  isWarningCommitteeName: boolean;
  isWarningConvenor: boolean;
  isWarningChairman: boolean;
  isWarningMeetingDate: boolean;
  isWarningMeetingSubject: boolean;
  isWarningMeetingMode: boolean;
  isWarningMeetingLink: boolean;

  isSmallScreen: boolean;
  isLoading:any;
}

const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("itemId");
  // const Id = params.get("itemId");

  return Id;
};

export interface IFileDetails {
  name?: string;
  content?: File;
  index?: number;
  fileUrl?: string;
  ServerRelativeUrl?: string;
  isExists?: boolean;
  Modified?: string;
  isSelected?: boolean;
}

export default class XenWpCommitteeMeetingsCreateForm extends React.Component<
  IXenWpCommitteeMeetingsFormsProps,
  CommtteeMeetingsStateProps
> {
  // private timerID: any;
  private _peopplePicker: IPeoplePickerContext;
  private _libraryName: any;
  private _listName: any;
  private _committeNameList: any;

  private _itemId: number = Number(getIdFromUrl());
  private _absUrl: any = this.props.context.pageContext.web.serverRelativeUrl;
  private _homePageUrl:any = this.props.homePageUrl

  // private currentDate: Date = new Date();
  // private _currentApprover:any;
  private _currentUserEmail = this.props.context.pageContext.user.email;

  constructor(props: any) {
    super(props);
    this.state = {
      committename: [],
      committeeNameFeildValue: "",
      charimanFeildValue: "",
      convernorFeildValue: "",
      convernorData: {},
      charimanData: {},
      selectedCommitteeMembers: [],
      selectedCommitteeGuestMembers: [],
      selectedCommitteeNoteRecords: [],
      committeeMembersData: [],
      committeeGuestMembersData: [],
      committeeNoteRecordsData: [],
      committeeMemberskey: [],
      committeeGuestMemberskey: [],
      committeeNoteRecordskey: [],
      CommitteeMeetingMemberCommentsDT:[],
      isModalOpen: false,
      isModalMOMOpen: false,
      modalMessage: "",
      meetingId: "To be generated",
      meetingDate: null,
      meetingSubject: "",
      meetingMode: "",
      meetingLink: "",
      statusNumber: "",
      MeetingStatus: "",
      auditTrail: [],
      committeeNoteRecordDropDownData: [],
      committeeNoteRecordSelectedValue: "",
      pdfLink: "",
      committeNoteRecordsDropDownDataWithAllProperties: [],
      dialogType: "",
      isMomDraftDialogOpen: false,
      draftResolutionFieldValue: "",
      selectedMOMNoteRecord: "",
      btnType: "",
      departmentAlias: "",
      invalidFields: [],

      isWarningCommitteeName: false,
      isWarningConvenor: false,
      isWarningChairman: false,
      isWarningMeetingDate: false,
      isWarningMeetingSubject: false,
      isWarningMeetingMode: false,
      isWarningMeetingLink: false,
      isLoading:true,

      isSmallScreen: window.innerWidth < 568,
    };

    this._peopplePicker = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };
    this.getfield();
    this._fetchDepartmentAlias();
    // console.log(this._homePageUrl)

    const listName = this.props.listName;
    this._listName = listName?.title;
    // console.log(this._listName, this.props.listName, "onload listName");

    const committeName = this.props.committeeMeetingNameList;
    this._committeNameList = committeName?.title;
    // console.log(
    //   this._committeNameList,
    //   this.props.committeeMeetingNameList,
    //   "onload committeeMeetingNameList"
    // );

    this._itemId && this._getItemData(this._itemId);

    const libraryTilte = this.props.libraryId;
    this._libraryName = libraryTilte?.title;

  }


  public componentDidMount() {
    // Add resize listener
    window.addEventListener('resize', this.handleResize);
    
    setTimeout(() => {
      this.setState({isLoading:false})
    }, 3000);
  }

  public componentWillUnmount() {
    // Remove resize listener
    window.removeEventListener('resize', this.handleResize);
  }

  private handleResize = () => {
    this.setState({ isSmallScreen: window.innerWidth < 768 });
  };

  private _getECommitteRequestBasedOnSelectedItem = async (
    committeeName: any
  ): Promise<any> => {
    try {
      // const fieldDetails = await this.props.sp.web.lists
      //   .getByTitle("CommitteeMeetingApprovers")();
      // console.log;

      const fieldDetails = await this.props.sp.web.lists
        .getByTitle("EcommiteeRequests")
        .items.select(
          "*",
          "CommitteeType",
          "Title",
          "Department",
          "CommitteeName"
        )
        .filter(
          `CommitteeType eq 'Committee' and CommitteeName eq '${committeeName}' and isMapped eq 'false'`
        )();
      // console.log(fieldDetails, "eCommittee Request ............");

      const dropDownDataListing = fieldDetails.filter(
        (each:any)=>{
          // console.log(each)
          return each.isMapped === false}
      ).map((each: any) => {
        return {
          key: each.Title,
          text: each.Title,
          id: each.Title,
        };
      });

      const committeNoteRecordsDropDownDataWithAllProperties =
        await Promise.all(
          fieldDetails.map(async (each: any) => {
            const link = await this._getItemDocumentsData(each.Title);

            return {
              key: each.Title,
              text: each.Title,
              id: each.Id,
              noteTitle: each.Title,
              committeeName: each.CommitteeName,
              department: each.Department,
              noteLink: link[0],
              link: link[1],
              mom: "",
            };
          })
        );

      this.setState({
        committeeNoteRecordDropDownData: dropDownDataListing,
        committeNoteRecordsDropDownDataWithAllProperties:
          committeNoteRecordsDropDownDataWithAllProperties,
      });

      // console.log(fieldDetails)

      // const profile = await this.props.sp.profiles.myProperties();

      // // this._userName = profile.DisplayName;
      // // this._role = profile.Title;

      // profile.UserProfileProperties.filter((element: any) => {
      //   if (element.Key === "Department") {
      //     this.setState({ department: element.Value });
      //   }
      // });

      // const filtering = fieldDetails.map((_x: { TypeDisplayName: string; InternalName: any; Choices: any; }) => {
      //   if (_x.TypeDisplayName === "Choice") {
      //     return [_x.InternalName, _x.Choices];
      //   }
      // });

      // Assuming fieldDetails is an array of items you want to add
      // this.setState((prevState) => ({
      //   itemsFromSpList: [...prevState.itemsFromSpList, ...finalList],
      //   isLoading: false,
      // }));
    } catch (error) {
      console.error("Error fetching field details: ", error);
    }
  };

  private _getFileObj = (data: any): any => {
    const tenantUrl = window.location.protocol + "//" + window.location.host;
    // console.log(tenantUrl);

    // const formatDateTime = (date: string | number | Date) => {
    //   const formattedDate = format(new Date(date), "dd-MMM-yyyy");
    //   const formattedTime = format(new Date(), "hh:mm a");
    //   return `${formattedDate} ${formattedTime}`;
    // };

    // const result = formatDateTime(data.TimeCreated);

    const result = data.TimeCreated;

    const filesObj = {
      name: data.Name,
      content: data,
      index: 0,
      fileUrl: tenantUrl + data.ServerRelativeUrl,
      ServerRelativeUrl: "",
      isExists: true,
      Modified: "",
      isSelected: false,
      size: parseInt(data.Length),
      type: `application/${data.Name.split(".")[1]}`,
      modifiedBy: data.Author.Title,
      createData: result,
    };
    // console.log(filesObj);
    return filesObj;
  };

  private _getItemDocumentsData = async (tilte: any): Promise<any> => {
    const folderName = tilte.replace(/\//g, "-");
    // console.log(folderName);

    const url = `${this._absUrl}/${this._libraryName}/${folderName}`;

    // console.log(url, "URL.........");

    try {
      const folderItemsPdf = await this.props.sp.web
        .getFolderByServerRelativePath(`${url}/Pdf`)
        .files.select("*")
        .expand("Author", "Editor")();

      // console.log(folderItemsPdf, "............", folderName);

      // console.log(folderItemsPdf);
      // console.log(folderItemsPdf[0]);
      // this.setState({noteTofiles:[folderItem]})
      let pdfLink;

      const tempFilesPdf: IFileDetails[] = [];
      folderItemsPdf.forEach((values: any) => {
        tempFilesPdf.push(this._getFileObj(values));
        pdfLink = this._getFileObj(values).fileUrl;
      });

      // console.log(pdfLink);
      // console.log(url)
      return [pdfLink, url];
    } catch {
      // console.log("failed to fetch");
    }
  };

  private _fetchDepartmentAlias = async (): Promise<void> => {
    try {
      // console.log("Starting to fetch department alias...");

      // Step 1: Fetch items from the Departments list
      const items: any[] = await this.props.sp.web.lists
        .getByTitle("Departments")
        .items.select(
          "Department",
          "DepartmentAlias",
          "Admin/EMail",
          "Admin/Title"
        ) // Fetching relevant fields
        .expand("Admin")();

      // console.log("Fetched items from Departments:", items);

      // let deparement = '';

      const profile = await this.props.sp.profiles.myProperties();

      // this._userName = profile.DisplayName;
      // this._role = profile.Title;

      profile.UserProfileProperties.filter((element: any) => {
        if (element.Key === "Department") {
          // department: element.Value

          const specificDepartment = items.find(
            (each: any) =>
              each.Department.includes(element.Value) ||
              each.Title?.includes(element.Value)
          );
    
          if (specificDepartment) {
            const departmentAlias = specificDepartment.DepartmentAlias;
            // console.log(
            //   "Department alias for department with 'Development' in title:",
            //   departmentAlias
            // );
    
            // Step 3: Update state with the department alias
            this.setState(
              {
                departmentAlias: departmentAlias, // Store the department alias
              },
              
            );
          }
        }
      });


      // Step 2: Find the department entry where the Title or Department contains "Development"
      
    } catch (error) {
      // console.error("Error fetching department alias: ", error);
    }
  };



  private _getTitle = (id:any):any=>{
    const currentyear = new Date().getFullYear();
    const nextYear = (currentyear + 1).toString().slice(-2);

    let meetingID = ''
     this.setState({
      meetingId: `${this.state.departmentAlias}/${currentyear}-${nextYear}/${id}`})

      meetingID =  `${this.state.departmentAlias}/${currentyear}-${nextYear}/${id}`
      // console.log(meetingID)

      return meetingID

  }


  

  private _getItemData = async (id: any) => {
    const item: any = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(id)
      .select(
        "*",

        "Chairman",
        "Chairman/Title",
        "Chairman/EMail",

        "CurrentApprover",

        "CurrentApprover/Title",

        "CurrentApprover/EMail"
      )
      .expand("Chairman", "CurrentApprover")();

    // console.log(`${id} ------Details`, item);

   

    this.setState({
      meetingId: item.Title,
      meetingDate: new Date(item.MeetingDate),
      meetingSubject: item.MeetingSubject,
      meetingMode: item.MeetingMode,
      meetingLink: item.MeetingLink,
      MeetingStatus: item.MeetingStatus,
      statusNumber: item.StatusNumber,
      committeeNameFeildValue: item.CommitteeName,
      committeeMembersData: JSON.parse(item.CommitteeMeetingMembersDTO),
      committeeGuestMembersData: JSON.parse(
        item.CommitteeMeetingGuestMembersDTO
      ),
      committeeNoteRecordsData: JSON.parse(item.CommitteeMeetingNoteDTO),
      CommitteeMeetingMemberCommentsDT:
      item.CommitteeMeetingMemberCommentsDT === null
        ? []
        : JSON.parse(item.CommitteeMeetingMemberCommentsDT),
      convernorFeildValue: item.Department,
      charimanFeildValue: item.Chairman.Title,
      auditTrail: JSON.parse(item.AuditTrail),
      isLoading:false
    });

    return item;
  };

  private getfield = async () => {
    try {
      // const fieldDetails = await this.props.sp.web.lists
      //   .getByTitle("CommitteeMeetingApprovers")();

      const fieldDetails = await this.props.sp.web.lists
        .getByTitle("CommitteeMeetingApprovers")
        .fields.filter("Hidden eq false and ReadOnlyField eq false")();

      // console.log(fieldDetails)

      // const profile = await this.props.sp.profiles.myProperties();

      // // this._userName = profile.DisplayName;
      // // this._role = profile.Title;

      // profile.UserProfileProperties.filter((element: any) => {
      //   if (element.Key === "Department") {
      //     this.setState({ department: element.Value });
      //   }
      // });

      const filtering = fieldDetails.map(
        (_x: { TypeDisplayName: string; InternalName: any; Choices: any }) => {
          if (_x.TypeDisplayName === "Choice") {
            return [_x.InternalName, _x.Choices];
          }
        }
      );

      const finalList = filtering?.filter((each: any) => {
        if (typeof each !== "undefined") {
          return each;
        }
      });

      finalList?.map((each: string | any[] | undefined) => {
        if (
          each !== undefined &&
          Array.isArray(each) &&
          each.length > 1 &&
          Array.isArray(each[1])
        ) {
          if (each[0] === "CommitteeName") {
            const committenameArray = each[1].map((item, index) => {
              return { key: item, text: item };
            });
            // console.log(committenameArray)

            this.setState({ committename: committenameArray });
          }
        }
      });

      // Assuming fieldDetails is an array of items you want to add
      // this.setState((prevState) => ({
      //   itemsFromSpList: [...prevState.itemsFromSpList, ...finalList],
      //   isLoading: false,
      // }));
    } catch (error) {
      console.error("Error fetching field details: ", error);
    }
  };

  private handleDeleteCommitteeMemberData = (index: number): void => {
    this.setState((prevState) => ({
      committeeMembersData: prevState.committeeMembersData.filter(
        (_: any, i: number) => i !== index
      ),
    }));
  };

  private columnsCommitteeMembers: IColumn[] = [
    {
      key: "sNo",
      name: "S.No",
      fieldName: "sNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "memberEmailName",
      name: "Members",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 230,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth:150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (_item: any, index?: number) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          title="Delete"
          ariaLabel="Delete"
          onClick={() => this.handleDeleteCommitteeMemberData(index!)}
        />
      ),
    },
  ];

  // private committeeMembersData = [
  //   {
  //     sNo: 1,
  //     members: "Dr. Alice Johnson",
  //     srNo: "CM001",
  //     designation: "Chairperson",
  //     action: "Active",
  //   },
  //   {
  //     sNo: 2,
  //     members: "Mr. Robert Lee",
  //     srNo: "CM002",
  //     designation: "Secretary",
  //     action: "Active",
  //   },
  //   {
  //     sNo: 3,
  //     members: "Ms. Sarah Connor",
  //     srNo: "CM003",
  //     designation: "Member",
  //     action: "Inactive",
  //   },
  //   {
  //     sNo: 4,
  //     members: "Mr. Tom Hardy",
  //     srNo: "CM004",
  //     designation: "Member",
  //     action: "Active",
  //   },
  // ];

  private handleDeleteGuestMember = (index: number): void => {
    this.setState((prevState) => ({
      committeeGuestMembersData: prevState.committeeGuestMembersData.filter(
        (_: any, i: number) => i !== index
      ),
    }));
  };

  private columnsCommitteeGuestMembers: IColumn[] = [
    {
      key: "sNo",
      name: "S.No",
      fieldName: "sNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "memberEmailName", // "guestMembers",
      name: "Guest Members",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 245,
      isResizable: true,
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (_item: any, index?: number) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          title="Delete"
          ariaLabel="Delete"
          onClick={() => this.handleDeleteGuestMember(index!)}
        />
      ),
    },
  ];

  // private committeeGuestMembersData = [
  //   {
  //     sNo: 1,
  //     guestMembers: "Dr. Alice Johnson",
  //     srNo: "001",
  //     designation: "Senior Advisor",
  //     action: "Confirmed",
  //   },
  //   {
  //     sNo: 2,
  //     guestMembers: "Mr. Bob Lee",
  //     srNo: "002",
  //     designation: "Technical Consultant",
  //     action: "Pending",
  //   },
  //   {
  //     sNo: 3,
  //     guestMembers: "Ms. Carla Davis",
  //     srNo: "003",
  //     designation: "Observer",
  //     action: "Declined",
  //   },
  //   {
  //     sNo: 4,
  //     guestMembers: "Dr. David Brown",
  //     srNo: "004",
  //     designation: "Specialist",
  //     action: "Invited",
  //   },
  // ];

  private _deleteRecord = (id: number) => {
    const updatedData = this.state.committeeNoteRecordsData.filter(
      (record: { id: number }) => record.id !== id
    );
    this.setState({ committeeNoteRecordsData: updatedData });
  };

  private columnsCommitteeNoteRecords: IColumn[] = [
    {
      key: "noteTitle",
      name: "Note Title",
      fieldName: "noteTitle",
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
    },
    {
      key: "committeeName",
      name: "Committee Name",
      fieldName: "committeeName",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "department",
      name: "Department",
      fieldName: "department",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "link",
      name: "Note Link",
      fieldName: "link",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
      onRender: (item) => {
        // Extract the file name from the URL

        return (
          <Link
            href={item.noteLink}
            download={item.noteTitle}
            // style={{ textDecoration: "none", color: "#0078D4" }}
          >
            {item.link}
          </Link>
        );
      },
    },
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item) =>
       ( this.state.statusNumber !== "1000" && this.state.statusNumber !== "2000" && this.state.statusNumber !== '') ? (
          <IconButton
            iconProps={{ iconName: "Edit" }}
            title="Edit"
            ariaLabel="Edit"
            onClick={() => {
              // console.log("Edit is Triggered");
              // console.log(item);
              this.setState({
                dialogType: "mom",
                isModalMOMOpen: true,
                selectedMOMNoteRecord: item.key,
                draftResolutionFieldValue:item.mom
              });
            }}
          />
        ) : (
          <IconButton
            iconProps={{ iconName: "Delete" }}
            title="Delete"
            ariaLabel="Delete"
            onClick={() => this._deleteRecord(item.id)}
          />
        ),
    },
  ];


  private columnsCommitteeComments: IColumn[] = [
    {
      key: "comments",
      name: "Comments",
      fieldName: "comments",
      minWidth: 200, // adjusted to match a percentage as close as possible
      maxWidth: 550,
      isResizable: true,
      // className: styles.columnHalf, // Apply the 50% width class
    },
    {
      key: "commentedBy",
      name: "Commented by",
      fieldName: "commentedBy",
      minWidth: 200,
      maxWidth: 550,
      isResizable: true,
      // className: styles.columnHalf, // Apply the 50% width class
    },
  ];

  // private committeeNoteRecordsData = [
  //   {
  //     noteTitle: "Budget Review",
  //     committeeName: "Finance Committee",
  //     department: "Finance",
  //     noteLink: "http://example.com/notes/budget-review",
  //     action: "Reviewed",
  //   },
  //   {
  //     noteTitle: "Hiring Policy Update",
  //     committeeName: "HR Committee",
  //     department: "Human Resources",
  //     noteLink: "http://example.com/notes/hiring-policy",
  //     action: "Approved",
  //   },
  //   {
  //     noteTitle: "Technology Upgrade",
  //     committeeName: "IT Committee",
  //     department: "Information Technology",
  //     noteLink: "http://example.com/notes/tech-upgrade",
  //     action: "Pending",
  //   },
  //   {
  //     noteTitle: "Marketing Strategy",
  //     committeeName: "Marketing Committee",
  //     department: "Marketing",
  //     noteLink: "http://example.com/notes/marketing-strategy",
  //     action: "In Progress",
  //   },
  // ];

  // public componentDidMount(): void {
  //   // Update the date every second
  //   this.timerID = setInterval(() => {
  //     this.currentDate = new Date();
  //     this.forceUpdate(); // Manually trigger a re-render
  //   }, 1000);
  // }

  // public componentWillUnmount(): void {
  //   if (this.timerID) {
  //     clearInterval(this.timerID); // Clear the interval on unmount
  //   }
  // }

  private _clearCommitteeMembersPeoplePicker = () => {
    this.setState({
      selectedCommitteeMembers: [],
      committeeMemberskey: this.state.committeeMemberskey + 1,
    }); // Update the key to force re-render
  };

  private _clearCommitteeGuestMembersPeoplePicker = () => {
    this.setState({
      selectedCommitteeGuestMembers: [],
      committeeGuestMemberskey: this.state.committeeGuestMemberskey + 1,
    }); // Update the key to force re-render
  };

  private _clearCommitteeNoteRecordsPeoplePicker = () => {
    this.setState({
      selectedCommitteeNoteRecords: [],
      committeeNoteRecordskey: this.state.committeeNoteRecordskey + 1,
    }); // Update the key to force re-render
  };

  //Committee Members Validation

  private checkSelectedCommitteMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember =
    (): boolean => {
      const committeeMemberEmail = this.state.committeeMembersData.map(
        (each: any) => each.email
      );

      const committeeGuestMembersEmail =
        this.state.committeeGuestMembersData.map((each: any) => each.email);

      const selectedCommitteeMembers = this.state.selectedCommitteeMembers[0];
      const selectedCommitteeMembersEmail =
        selectedCommitteeMembers.email ||
        selectedCommitteeMembers.secondaryText;
      const selectedMemberIsAChairman =
        selectedCommitteeMembersEmail === this.state.charimanData.EMail;
      // console.log(selectedMemberIsAChairman);

      const selectedMemberIsACon =
      selectedCommitteeMembersEmail === this.state.convernorData.EMail;
    // console.log(selectedMemberIsACon);

      // Condition checks
      const iscommitteeMemberOrcommitteeGuestMembers =
        committeeGuestMembersEmail.includes(selectedCommitteeMembersEmail) ||
        committeeMemberEmail.includes(selectedCommitteeMembersEmail);

      const isCurrentUserCommitteMember =
        this._currentUserEmail === selectedCommitteeMembersEmail;

      return (
        iscommitteeMemberOrcommitteeGuestMembers ||
        isCurrentUserCommitteMember ||
        selectedMemberIsAChairman ||selectedMemberIsACon
      );
    };

  //Guest Members Validation

  private checkSelectedGuestMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember =
    (): boolean => {
      const committeeMemberEmail = this.state.committeeMembersData.map(
        (each: any) => each.email
      );

      const committeeGuestMembersEmail =
        this.state.committeeGuestMembersData.map((each: any) => each.email);

      const selectedGuestMembers = this.state.selectedCommitteeGuestMembers[0];
      const selectedGuestMembersEmail =
        selectedGuestMembers.email || selectedGuestMembers.secondaryText;

      const selectedMemberIsAChairman =
        selectedGuestMembersEmail === this.state.charimanData.EMail;
      // console.log(selectedMemberIsAChairman);

      // Condition checks
      const iscommitteeMemberOrcommitteeGuestMembers =
        committeeGuestMembersEmail.includes(selectedGuestMembersEmail) ||
        committeeMemberEmail.includes(selectedGuestMembersEmail);

      const isCurrentUserGuestMember =
        this._currentUserEmail === selectedGuestMembersEmail;

      return (
        iscommitteeMemberOrcommitteeGuestMembers ||
        isCurrentUserGuestMember ||
        selectedMemberIsAChairman
      );
    };

  private handleOnAdd = async (event: any, type: string): Promise<void> => {
    if (this.state.committeeNameFeildValue === "") {
      this.setState({
        isModalOpen: true,
        modalMessage: "Please select Committee Name and click Add.",
      });
      this._clearCommitteeMembersPeoplePicker();
      return;
    }

    if (type === "committeeMembers") {
      if (
        this.checkSelectedCommitteMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember()
      ) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "The selected member cannont be same as existing Member/Convenor/Chairman.",
        });
        this._clearCommitteeMembersPeoplePicker();
        this._clearCommitteeGuestMembersPeoplePicker();
        return;
      }

      this.setState({
        committeeMembersData: [
          ...this.state.committeeMembersData,
          ...this.state.selectedCommitteeMembers,
        ],
      });
      this._clearCommitteeMembersPeoplePicker();
    } else if (type === "committeeGuestMembers") {
      if (
        this.checkSelectedGuestMemberIsInRequestorOrChairmanOrGuestMemberOrCommitteMember()
      ) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "The selected member cannont be same as existing Member/Convenor/Chairman.",
        });
        this._clearCommitteeMembersPeoplePicker();
        this._clearCommitteeGuestMembersPeoplePicker();
        return;
      }
      this.setState({
        committeeGuestMembersData: [
          ...this.state.committeeGuestMembersData,
          ...this.state.selectedCommitteeGuestMembers,
        ],
      });
      this._clearCommitteeGuestMembersPeoplePicker();
    } else if (type === "committeeNoteRecords") {
      this.setState({
        committeeNoteRecordsData: [
          ...this.state.committeeNoteRecordsData,
          ...this.state.selectedCommitteeNoteRecords,
        ],
      });
      this._clearCommitteeNoteRecordsPeoplePicker();
    }
  };

  // private getFormattedDate = (): string => {

  //   return `${this.currentDate.getDate()}-${
  //     this.currentDate.getMonth() + 1
  //   }-${this.currentDate.getFullYear()} ${this.currentDate.getHours()}:${this.currentDate.getMinutes()}:${this.currentDate.getSeconds()}`;
  // };

  private _getUserProperties = async (loginName: any): Promise<any> => {
    // console.log(loginName)
    let designation = "NA";
    let email = "NA";
    // const loginName = this.state.peoplePickerData[0]
    const profile = await this.props.sp.profiles.getPropertiesFor(loginName);
    // console.log(profile)
    // console.log(profile.DisplayName);
    // console.log(profile.Email);
    // console.log(profile.Title);
    // console.log(profile.UserProfileProperties.length);
    designation = profile.Title;
    email = profile.Email;
    // Properties are stored in inconvenient Key/Value pairs,
    // so parse into an object called userProperties
    const props: any = {};
    profile.UserProfileProperties.forEach(
      (prop: { Key: string | number; Value: any }) => {
        props[prop.Key] = prop.Value;
      }
    );

    profile.userProperties = props;
    // console.log("Account Name: " + profile.userProperties.AccountName);
    return [designation, email];
  };

  private _getPeoplePickerItemsCommitteeMembers = async (
    items: any[]
  ): Promise<any> => {
   
   
   
    // console.log("Items:", items);
    // console.log(this.props.typeOFButton)
    // console.log("Items:", items);
    // fetchedData = items
    // console.log(items[0].loginName);

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    // console.log(items, "this._getUserProperties(items[0].loginName)");

    // this.setState({approverInfo:items})

    const dataRec = await this._getUserProperties(items[0].loginName);
    // const finalData = await dataRec.json()
    // dataRec.then((x: any)=>{
    //   console.log(x)
    //   designation=x
    // });
    // console.log(typeof dataRec?.toString());

    if (typeof dataRec[0]?.toString() === "undefined") {
      const newItemsDataNA = items.map((obj: any) => {
        // console.log(obj);
        return {
          ...obj,
          optionalText: "N/A",

          email: obj.secondaryText,
        };
      });
      // console.log(newItemsDataNA);
      this.setState({
        selectedCommitteeMembers: [
          {
            memberEmailName: newItemsDataNA[0].text,
            srNo: newItemsDataNA[0].srNo,
            designation: newItemsDataNA[0].optionalText,
            email: newItemsDataNA[0].email,
            userId: newItemsDataNA[0].id,
          },
        ],
      });
    } else {
      const newItemsData = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: dataRec[0],

          email: dataRec[1],
          srNo: dataRec[1].split("@")[0],
        };
      });
      // console.log(newItemsData)
      // this.props.getDetails(newItemsData,this.props.typeOFButton)
      // // eslint-disable-next-line no-unused-expressions
      // newItemsData.length > 0 && this.props.clearPeoplePicker(this._clearPeoplePicker,"clearFuntion")
      this.setState({
        selectedCommitteeMembers: [
          {
            memberEmailName: newItemsData[0].text,
            srNo: newItemsData[0].srNo,
            designation: newItemsData[0].optionalText,
            email: newItemsData[0].email,
            userId: newItemsData[0].id,
          },
        ],
      });
      // this._clearPeoplePicker();
    }
  };

  private _getPeoplePickerItemsCommitteeGuestMembers = async (
    items: any[]
  ): Promise<any> => {
    // console.log("Items:", items);
    // console.log(this.props.typeOFButton)
    // console.log("Items:", items);
    // fetchedData = items
    // console.log(items[0].loginName);

    // eslint-disable-next-line @typescript-eslint/no-floating-promises
    // console.log(items, "this._getUserProperties(items[0].loginName)");

    // this.setState({approverInfo:items})

    const dataRec = await this._getUserProperties(items[0].loginName);
    // const finalData = await dataRec.json()
    // dataRec.then((x: any)=>{
    //   console.log(x)
    //   designation=x
    // });
    // console.log(typeof dataRec?.toString());

    if (typeof dataRec[0]?.toString() === "undefined") {
      const newItemsDataNA = items.map((obj: any) => {
        // console.log(obj);
        return {
          ...obj,
          optionalText: "N/A",

          email: obj.secondaryText,
        };
      });
      // console.log(newItemsDataNA);
      this.setState({
        selectedCommitteeGuestMembers: [
          {
            memberEmailName: newItemsDataNA[0].text,
            srNo: newItemsDataNA[0].srNo,
            designation: newItemsDataNA[0].optionalText,
            email: newItemsDataNA[0].email,
            userId: newItemsDataNA[0].id,
          },
        ],
      });
    } else {
      const newItemsData = items.map((obj: any) => {
        return {
          ...obj,
          optionalText: dataRec[0],

          email: dataRec[1],
          srNo: dataRec[1].split("@")[0],
        };
      });
      // // console.log(newItemsData)
      // this.props.getDetails(newItemsData,this.props.typeOFButton)
      // // eslint-disable-next-line no-unused-expressions
      // newItemsData.length > 0 && this.props.clearPeoplePicker(this._clearPeoplePicker,"clearFuntion")
      this.setState({
        selectedCommitteeGuestMembers: [
          {
            memberEmailName: newItemsData[0].text,
            srNo: newItemsData[0].srNo,
            designation: newItemsData[0].optionalText,
            email: newItemsData[0].email,
            userId: newItemsData[0].id,
          },
        ],
      });
      // this._clearPeoplePicker();
    }
  };

  // private _getPeoplePickerItemsCommitteeNoteRecords =async (items: any[]): Promise<any> => {
  //   console.log("Items:", items);
  //   // console.log(this.props.typeOFButton)
  //   // console.log("Items:", items);
  //   // fetchedData = items
  //   // console.log(items[0].loginName);

  //   // eslint-disable-next-line @typescript-eslint/no-floating-promises
  //   // console.log(items, "this._getUserProperties(items[0].loginName)");

  //   // this.setState({approverInfo:items})

  //   const dataRec = await this._getUserProperties(items[0].loginName);
  //   // const finalData = await dataRec.json()
  //   // dataRec.then((x: any)=>{
  //   //   console.log(x)
  //   //   designation=x
  //   // });
  //   // console.log(typeof dataRec?.toString());

  //   if (typeof dataRec[0]?.toString() === "undefined") {
  //     const newItemsDataNA = items.map(
  //       (obj:any) => {
  //         // console.log(obj);
  //         return {
  //           ...obj,
  //           optionalText: "N/A",

  //           email: obj.secondaryText,
  //         };
  //       }
  //     );
  //     // console.log(newItemsDataNA);
  //     this.setState({ selectedCommitteeNoteRecords: [{
  //       member:newItemsDataNA[0].text,
  //       srNo:newItemsDataNA[0].srNo,
  //       designation:newItemsDataNA[0].optionalText,email:newItemsDataNA[0].email, userId:newItemsDataNA[0].id

  //     }] });
  //   } else {
  //     const newItemsData = items.map((obj: any) => {
  //       return {
  //         ...obj,
  //         optionalText: dataRec[0],

  //         email: dataRec[1],
  //         srNo: dataRec[1].split("@")[0],
  //       };
  //     });
  //     // console.log(newItemsData)

  //     this.setState({ selectedCommitteeNoteRecords: [{
  //       member:newItemsData[0].text,
  //       srNo:newItemsData[0].srNo,
  //       designation:newItemsData[0].optionalText
  //       ,email:newItemsData[0].email,
  //       userId:newItemsData[0].id

  //     }] });
  //     // this._clearPeoplePicker();
  //   }
  // };

  private onRenderCaretDowncommitteeNameFeildValue = (): JSX.Element => {
    return this.state.committeeNameFeildValue ? (
      <Icon
        iconName="Cancel"
        onClick={() => {
          this.setState({ committeeNameFeildValue: "" });
        }}
      />
    ) : (
      <Icon
        iconName="ChevronDown"
        onClick={() => {
          this.setState({ committeeNameFeildValue: "" });
        }}
      />
    );
  };

  private handleCommittename = async (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): Promise<void> => {
    const value = option ? option.text : "";

    if (
      value === "IBCommitte" ||
      value === "AccountentCommittee" ||
      value === "Icc Committee"
    ) {
      const fieldDetails = await this.props.sp.web.lists
        .getByTitle(this._committeNameList)
        .items.select(
          "*",
          "Convener",

          "Convener/Title",
          "Chairman",
          "Chairman/Title",
          "Chairman/EMail",
          "Convener/EMail"
        )
        .expand("Convener", "Chairman")();

       
      // console.log(fieldDetails);
      fieldDetails.filter(async (each: any) => {
        if (each.CommitteeNames === value) {
          // console.log(each,"committe Name List")
          const convenorDepartment = await this.getUserDepartmentByEmail(each.ConvenerId)
          // console.log(convenorDepartment)
          this.setState({
            charimanFeildValue: each.Chairman.Title,
            // convernorFeildValue: each.Department,
            convernorData: { ...each.Convener, convenerId: each.ConvenerId,department:convenorDepartment },
            charimanData: { ...each.Chairman, chairmanId: each.ChairmanId },
          });
        }
      });
    }
    this._getECommitteRequestBasedOnSelectedItem(value);

    this.setState({
      committeeNameFeildValue: value,
      committeeGuestMembersData: [],
      committeeMembersData: [],
      committeeNoteRecordsData: [],
    });
  };

  private onDateChange = (date: any): void => {
    this.setState({ meetingDate: date }); // Update the state with the selected date, or null if the user clears the selection
    // console.log("Selected date:", date);
  };

  private _closeModal = (): void => {
    this.setState({ isModalOpen: false ,isModalMOMOpen:false});
    this.setState({ dialogType: "" });
  };

  private _getCommitteeMeetingMembersDTO = (data: any): any => {
    const makeCommitteeMeetingMemberDTO = data.map((each: any, index: any) => {
      // console.log(each, "checking for srNo");

      return {
        createdDate: new Date(),
        userId: each.userId,
        srNo: each.srNo,
        designation: each.designation,
        memberEmail: each.email||each.memberEmail,
        memberEmailName: each.memberEmailName,
        status: index === 0 ? "Pending" : "Waiting",
        statusNumber: "",
      };
    });

    // console.log(makeCommitteeMeetingMemberDTO);
    return JSON.stringify(makeCommitteeMeetingMemberDTO);
  };

  private _getMemeberId = (data: any) => {
    const getId = data.map((each: any) => each.userId);
    // console.log(getId);
    return getId;
  };

  private _getCurrrentApproverId = (data: any): any => {
    const currentApproverId = data.filter((each: any, index: any) => {
      if (index === 0) {
        return each;
      }
    });

    // console.log(currentApproverId[0].userId);
    return currentApproverId[0].userId;
  };

  private createEcommitteeMeetingObject = async (
    status: string,
    statusNumber: any
  ): Promise<any> => {
    // console.log(status);

    const auditTrail = this.state.auditTrail;

    const auditLog = {
      actionBy: this.props.context.pageContext.user.displayName,

      action: `Committee Meeting ${status}`,

      
      actionDate: new Date().toLocaleDateString(),
    };
    auditTrail.push(auditLog);

    const ecommitteObject: any = {
      MeetingDate: this.state.meetingDate,
      MeetingLink: this.state.meetingLink,

      MeetingMode: this.state.meetingMode,
      MeetingSubject: this.state.meetingSubject,
      CommitteeName: this.state.committeeNameFeildValue,
      ChairmanId: this.state.charimanData.chairmanId,
      FinalApproverId:
        this.state.committeeMembersData[
          this.state.committeeMembersData.length - 1
        ].userId,
      CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
        this.state.committeeMembersData
      ),
      CommitteeMeetingGuestMembersDTO: this._getCommitteeMeetingMembersDTO(
        this.state.committeeGuestMembersData
      ),
      CommitteeMeetingNoteDTO: JSON.stringify(
        this.state.committeeNoteRecordsData
      ),
      CurrentApproverId: this._getCurrrentApproverId(
        this.state.committeeMembersData
      ),
      CommitteeMeetingMembersId: this._getMemeberId(
        this.state.committeeMembersData
      ),
      CommitteeMeetingGuestsId: this._getMemeberId(
        this.state.committeeGuestMembersData
      ),
      MeetingStatus: status,
      ConvenerDTO:JSON.stringify(this.state.convernorData),
      Department: this.state.convernorFeildValue,
      StatusNumber: statusNumber,
      AuditTrail: JSON.stringify(auditTrail),
      startProcessing: true,
      PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id
    };
    // console.log(ecommitteObject);
    return ecommitteObject;
  };

  private _handleCreateMeeting = async (): Promise<void> => {
    this.setState({ isModalOpen: false,isLoading:true });


      
      await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.add(await this.createEcommitteeMeetingObject("Created", "1000")).then(
        async (res:any)=>{
          const title = this._getTitle(res.Id)
          // console.log(title)

          await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(res.Id)
          .update({
            Title: title,
           
          })

        }
      );
    // const id = response.Id;
    // console.log(id);

    




   
    this.setState({isLoading:false, isModalOpen: true, dialogType: "success" });
  };

  private _handlePulbicMeeting = async (): Promise<void> => {
    this.setState({ isModalOpen: false,isLoading:true });
    
    const auditTrail = {
      action: `Committee Meeting Published`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    }

    

      this._itemId
      ? await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(this._itemId)
          .update({
            StatusNumber: "2000",
            MeetingStatus: "Published",

            AuditTrail: JSON.stringify([...this.state.auditTrail,auditTrail]),
            startProcessing: true,
            PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id
          })
      : await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.add(
            await this.createEcommitteeMeetingObject("Published", "2000")
          ).then(
            async (res:any)=>{
              const title = this._getTitle(res.Id)
              // console.log(title)
    
              await this.props.sp.web.lists
              .getByTitle(this._listName)
              .items.getById(res.Id)
              .update({
                Title: title,
               
              })
    
            }
          );
    // const id = response.Id;
    // console.log(id);
    this.setState({isLoading:false, isModalOpen: true, dialogType: "success" });
  };


  private _handleReturnBack = async (): Promise<void> => {
    // console.log("Returned back triggered")
    this.setState({ isModalOpen: false,isLoading:true });
    const auditTrail = this.state.auditTrail || [];
    // const comments = this.state.CommitteeMeetingMemberCommentsDT;

   
    // console.log(comments)


    auditTrail.push({
      action: `Meeting Returned Back`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });

      
       await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(this._itemId)
          .update(
            {
              CommitteeMeetingMembersDTO: this._getCommitteeMeetingMembersDTO(
                this.state.committeeMembersData
              ),

              CurrentApproverId: this._getCurrrentApproverId(
                this.state.committeeMembersData
              ),
              
              MeetingStatus: "Return Back",
              StatusNumber: "5000",
              AuditTrail: JSON.stringify(auditTrail),
              startProcessing: true,

            }
          )

          await this.props.sp.web.lists
          .getByTitle(this._listName)
          .items.getById(this._itemId)
          .update(
            {
             MeetingStatus: "Pending Approval",
              StatusNumber: "5000",
              AuditTrail: JSON.stringify(auditTrail),

            }
          )


      
    // const id = response.Id;
    // console.log(id);
    this.setState({isLoading:false, isModalOpen: true, dialogType: "success" });
  };

  private _handleMeetingOver = async (): Promise<void> => {
    this.setState({ isModalOpen: false,isLoading:true });
    const auditTrail = this.state.auditTrail || [];

    auditTrail.push({
      action: `Meeting Over`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });

     await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "3000",
        AuditTrail: JSON.stringify(auditTrail),
        startProcessing: true,
        MeetingStatus: "Meeting Over",
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id
      });
    // const id = response.Id;
   
    // console.log(id);
    this.setState({isLoading:false, isModalOpen: true, dialogType: "success" });
  };

  private _handleMOMPublished = async (): Promise<void> => {
    this.setState({ isModalOpen: false,isLoading:true });
    const auditTrail = this.state.auditTrail || [];

    auditTrail.push({
      action: `Meeting MOM Published`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });

     await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "4000",

        MeetingStatus: "MOM Published",
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingNoteDTO: JSON.stringify(
          this.state.committeeNoteRecordsData
        ),
        startProcessing: true,
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id
      });

    await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(this._itemId)
      .update({
        StatusNumber: "5000",
        MeetingStatus: "Pending Approval",

        startProcessing: true,
      });

    await this.state.committeeNoteRecordsData.map(async (each: any) => {
      await this.props.sp.web.lists
        .getByTitle("EcommiteeRequests")
        .items.getById(each.id)
        .update({
          isMapped: true,
        });
    });
    // const id = response.Id;
    // console.log(id);
    this.setState({isLoading:false, isModalOpen: true, dialogType: "success" });
  };

  private options: IDropdownOption[] = [
    { key: "teams", text: "Teams" },
    { key: "webex", text: "WebEx" },
  ];

  // Event handler for the dropdown
  private handleMeetingModeChange = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ meetingMode: option.key });
      // console.log(`Selected meeting mode: ${option.text}`);
    }
  };

  // Event handler for handling changes in both textareas
  private handleInputChange = (
    event: React.ChangeEvent<HTMLTextAreaElement>,
    field: "meetingSubject" | "meetingLink"
  ): void => {
    const { value } = event.target;
    this.setState({ [field]: value } as Pick<
      CommtteeMeetingsStateProps,
      "meetingSubject" | "meetingLink"
    >);
  };

  private columnsCommitteeWorkFlowLog: IColumn[] = [
    {
      key: "action",
      name: "Action",
      fieldName: "action",
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "actionBy",
      name: "Action By",
      fieldName: "actionBy",
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: "actionDate",
      name: "Action Date",
      fieldName: "actionDate",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  private handleCommitteeNoteRecordsDropdownChange = (
    event: any,
    option: any
  ) => {
    this.setState({ committeeNoteRecordSelectedValue: option.key });
  };

  private _handleOnAddCommitteeNoteRecords = (): any => {
    this.state.committeeNoteRecordDropDownData.filter(
      (each: any) => each.key === this.state.committeeNoteRecordSelectedValue
    );

    this.setState({
      committeeNoteRecordsData: [
        ...this.state.committeeNoteRecordsData,
        ...this.state.committeNoteRecordsDropDownDataWithAllProperties.filter(
          (each: any) =>
            each.key === this.state.committeeNoteRecordSelectedValue
        ),
      ],
    });
  };

  private _checkAllNoteRecord = (): boolean => {
    const isAnyMOMEmpty = this.state.committeeNoteRecordsData.some(
      (each: any) => {
        // console.log(each, "Each Note Record");
        // console.log(each.mom === "", "each MOM is empty...");
        return each.mom === "";
      }
    );
  
    // console.log(isAnyMOMEmpty, "Is any MOM empty?");
    return isAnyMOMEmpty;
  };
  
  private stylesModal = mergeStyleSets({
    modal: {
      padding: "10px",
      minWidth: "300px",
      maxWidth: "80vw",
      width: "100%",
      "@media (min-width: 768px)": {
        maxWidth: "580px", // Adjust width for medium screens
      },
      "@media (max-width: 767px)": {
        maxWidth: "290px", // Adjust width for smaller screens
      },
      margin: "auto",
      backgroundColor: "white",
      borderRadius: "4px",
      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
    },
    header: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      borderBottom: "1px solid #ddd",
      minHeight: "50px",
    },
    headerTitle: {
      margin: "5px",
      marginLeft: "5px",
      fontSize: "16px",
      fontWeight: "400",
    },
    peoplePickerAndAddCombo: {
      display: "flex",
      gap: "5px",
      width: "60%",
    },
    body: {
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      textAlign: "center",
      padding: "20px 0",
    },
    footer: {
      display: "flex",
      justifyContent: "flex-end",
      marginTop: "20px",
      borderTop: "1px solid #ddd", // Added border to the top of the footer
      paddingTop: "10px",
    },
  });

  private stylesMOMModal = mergeStyleSets({
    modal: {
     
      // width: "100%",
      // height:'350px',
      padding: '10px',
      // width:'1000px',
      width:'580px',
      
      "@media (min-width: 768px)": {
        maxWidth: "580px", // Adjust width for medium screens
      },
      "@media (max-width: 767px)": {
        maxWidth: "290px", // Adjust width for smaller screens
      },
     
      backgroundColor: "white",
      borderRadius: "4px",
      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
    },
    header: {
      display: "flex",
      justifyContent: "space-between",
      alignItems: "center",
      borderBottom: "1px solid #ddd",
      minHeight: "50px",
    },
    headerTitle: {
      margin: "5px",
      marginLeft: "5px",
      fontSize: "16px",
      fontWeight: "400",
    },
    body: {
      marginTop:'48px',
      position: 'relative',
      height:'150px'
    },
    // richTextWrapper: {
    //   width: "90%", // Matches modal width
    // },
    footer: {
      display: "flex",
      justifyContent: "flex-end",
      marginTop: "20px",
      borderTop: "1px solid #ddd", // Added border to the top of the footer
      padding: "10px",
    },
  });

  private _getAlertDialogContent = (): any => {
    return (
      <>
        <div className={this.stylesModal.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={this.stylesModal.headerTitle}>Alert</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={() => {
              // console.log("Triggered close");
              this._closeModal();
            }}
          />
        </div>
        <div className={this.stylesModal.body}>
          <p>{this.state.modalMessage}</p>
        </div>
        <div className={this.stylesModal.footer}>
          <PrimaryButton
            iconProps={{ iconName: "ReplyMirrored" }}
            // onClick={this._closeModal}

            onClick={() => {
              // if (this.state.warnType !=="no"){
              // //   const pageURL: string = this.props.homePageUrl;
              // // window.location.href = `${pageURL}`;

              // }
              this._closeModal();
            }}
            text="OK"
          />
        </div>
      </>
    );
  };

  private _onRichTextChangeForMom = (newText: string) => {
    // this.properties.myRichText = newText;
    // console.log(newText);
    this.setState({ draftResolutionFieldValue: newText });
    return newText;
  };

  // private _getMOMtDialogContent = (): any => {
  //   return (
  //     <>
  //     <div className={this.stylesModal.header}>
  //       <div style={{ display: "flex", alignItems: "center" }}>
  //         <IconButton iconProps={{ iconName: "Info" }} />
  //         <h4 className={this.stylesModal.headerTitle}>ADD MOM</h4>
  //       </div>
  //       <IconButton
  //         iconProps={{ iconName: "Cancel" }}
  //         onClick={() => {
  //           this._closeModal();
  //           this.setState({ dialogType: "" });
  //         }}
  //       />
  //     </div>
  //     <div className={this.stylesModal.body}>
  //     <RichText
  //             value={this.state.draftResolutionFieldValue}
  //             onChange={(text: string) => this._onRichTextChangeForMom(text)}

  //           />
  //     </div>
  //     <div className={this.stylesModal.footer}>
  //       <PrimaryButton
  //         iconProps={{ iconName: "Add" }}
  //         onClick={() => {
  //           const updatedData = this.state.committeeNoteRecordsData.map((each: any) => {
  //             if (each.key === this.state.selectedMOMNoteRecord) {
  //               return { ...each, mom: this.state.draftResolutionFieldValue };
  //             }
  //             return each;
  //           });

  //           this.setState({
  //             draftResolutionFieldValue: "",
  //             committeeNoteRecordsData: updatedData,
  //             dialogType: "",
  //           });

  //           this._closeModal();
  //         }}
  //         text="Add"
  //       />
  //     </div>
  //   </>

  //   );
  // };

  private _getConfirmationtDialogContent = (): any => {
    const styles = mergeStyleSets({
      modal: {
        minWidth: "300px",
        maxWidth: "80vw",
        width: "100%",
        "@media (min-width: 768px)": {
          maxWidth: "580px", // Adjust width for medium screens
        },
        "@media (max-width: 767px)": {
          maxWidth: "290px", // Adjust width for smaller screens
        },
        margin: "auto",
        padding: "10px",
        backgroundColor: "white",
        borderRadius: "4px",
        // height:'260px',
        // display:'flex',
        // flexDirection:'column',
        // alignItem:'center',
        // justifyContent:'center',

        boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
      },
      header: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        borderBottom: "1px solid #ddd",
        minHeight: "50px",
        padding: "5px",
      },
      headerTitle: {
        margin: "5px",
        marginLeft: "5px",
        fontSize: "16px",
        fontWeight: "400",
      },
      headerIcon: {
        paddingRight: "0px", // Reduced space between the icon and the title
      },
      body: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        textAlign: "center",
        padding: "20px 0",
        height: "100%",
        "@media (min-width: 768px)": {
          marginLeft: "20px", // Adjust width for smaller screens
          marginRight: "20px", // Adjust width for medium screens
        },
        "@media (max-width: 767px)": {
          marginLeft: "20px", // Adjust width for smaller screens
          marginRight: "20px",
        },
      },
      footer: {
        display: "flex",
        justifyContent: "space-between", // Adjusted to space between

        borderTop: "1px solid #ddd",
        paddingTop: "10px",
        minHeight: "50px",
      },
      button: {
        maxHeight: "32px",
        flex: "1 1 50%", // Ensures each button takes up 50% of the footer width
        margin: "0 5px", // Adds some space between the buttons
      },
      buttonContent: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      },
      buttonIcon: {
        marginRight: "4px", // Adjust the space between the icon and text
      },

      removeTopMargin: {
        marginTop: "4px",
        marginBottom: "14px",
        fontWeight: "400",
      },
    });
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton
              iconProps={{ iconName: "WaitlistConfirm" }}
              className={styles.headerIcon}
            />
            <h4 className={styles.headerTitle}>Confirmation</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={this._closeModal}
          />
        </div>
        <div className={styles.body}>
          <p className={`${styles.removeTopMargin}`}>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Are you sure you want to Create this Meeting?";
                case "publishMeeting":
                  return "Are you sure you want to Publish this Meeting?";
                case "meetingOver":
                  return "Are you sure you want to submit meeting over?";
                case "momPublished":
                  return "Are you sure you want to Publish MOM?";

                  case "returnBack":
                    return "Are you sure you want to Return Back the Meeting?";
              }
            })()}
          </p>
          <br/>
          <p className={`${styles.removeTopMargin}`}>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Please check the details filled  and click on Confirm button to Create meeting.";
                case "publishMeeting":
                  return "Please check the details filled  and click on Confirm button to Publish meeting.";
                case "meetingOver":
                  return "Please check the details filled  and click on Confirm button to submit meeting over.";
                case "momPublished":
                  return "Please check the details filled  and click on Confirm button to Publish MOM.";
              
                  case "returnBack":
                    return "Please check the details filled  and click on Confirm button to Return Back the meeting."; 

                }
            })()}
          </p>
        </div>
        <div className={styles.footer}>
          <PrimaryButton
            onClick={() => {
              switch (this.state.btnType) {
                case "momPublished":
                  return this._handleMOMPublished();
                case "meetingOver":
                  return this._handleMeetingOver();
                case "publishMeeting":
                  return this._handlePulbicMeeting();

                  case "returnBack":
                    return this._handleReturnBack();
                default:
                  return this._handleCreateMeeting();
              }
            }}
            text="Confirm"
            iconProps={{
              iconName: "SkypeCircleCheck",
              styles: { root: styles.buttonIcon },
            }}
            styles={{ root: styles.buttonContent }}
            className={styles.button}
          />
          <DefaultButton
            onClick={this._closeModal}
            text="Cancel"
            iconProps={{
              iconName: "ErrorBadge",
              styles: { root: styles.buttonIcon },
            }}
            styles={{ root: styles.buttonContent }}
            className={styles.button}
          />
        </div>
      </>
    );
  };

  private _geSuccessDialogContent = (): any => {
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={styles.headerTitle}>Alert</h4>
          </div>
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={this._closeModal}
          />
        </div>
        <div className={styles.body}>
          <p>
            {(() => {
              switch (this.state.btnType) {
                case "createMeeting":
                  return "Meeting has been created successfully!";
                case "publishMeeting":
                  return "Meeting has been published successfully!";
                case "meetingOver":
                  return "Meeting over has been submitted successfully!";
                case "momPublished":
                  return "Meeting minutes has been published successfully!";

                  case "returnBack":
                    return "Return back has been successfully!";
              }
            })()}
          </p>
          {/* {statusOfReq === 'approver changed'?<p>The current actioner(Approver/Reviewer/Referee) has been updated successfully.</p>:<p>The request for {typeOfNote} note has been {statusOfReq.toLowerCase()} successfully.</p>} */}
        </div>
        <div className={styles.footer}>
          <PrimaryButton
            className={styles.button}
            iconProps={{ iconName: "ReplyMirrored" }}
            onClick={() => {
              this._closeModal();
              const pageURL: string = this.props.homePageUrl;
              // console.log(pageURL)
              window.location.href = `${pageURL}`;
            }}
            text="OK"
          />
        </div>
      </>
    );
  };

  private _getValidationDialog = (): any => {
    return (
      <>
        <div className={styles.header}>
          <div style={{ display: "flex", alignItems: "center" }}>
            <IconButton iconProps={{ iconName: "Info" }} />
            <h4 className={styles.headerTitle}>Alert</h4>
          </div>
          <Icon iconName="Cancel" onClick={this._closeModal} />
        </div>
        <div className={styles.body} style={{ alignItems: "flex-start" }}>
          <h4 className={styles.headerTitle}>
            Please fill up all the mandatory fields
          </h4>
          <ul>
            {this.state.invalidFields.length > 0 &&
              this.state.invalidFields.map((each: any) => (
                <li style={{ textAlign: "left" }} key={each}>
                  {each}
                </li>
              ))}
          </ul>
        </div>

        <div className={styles.footer}>
          <PrimaryButton
            text="OK"
            iconProps={{ iconName: "ReplyMirrored" }}
            onClick={this._closeModal}
            // styles={buttonStyles}
          />
        </div>
      </>
    );
  };

  private _getBodyContentOfDialogBox = (): any => {
    switch (this.state.dialogType) {
      // case "mom":
      //   return

      case "confirmation":
        return this._getConfirmationtDialogContent();


      case "success":
        return this._geSuccessDialogContent();

      case "validation":
        return this._getValidationDialog();
      default:
        return this._getAlertDialogContent();
    }
  };

  private _checkFields = (): boolean => {
    const fieldValues: any = {
      committeeName: [
        this.state.committeeNameFeildValue,
        "Committee Name",
        "isWarningCommitteeName",
      ],
      convenor: [
        this.state.convernorFeildValue,
        "Convenor",
        "isWarningConvenor",
      ],
      chairman: [
        this.state.charimanFeildValue,
        "Chairman",
        "isWarningChairman",
      ],
      meetingDate: [
        this.state.meetingDate,
        "Meeting Date",
        "isWarningMeetingDate",
      ],
      meetingSubject: [
        this.state.meetingSubject,
        "Meeting Subject",
        "isWarningMeetingSubject",
      ],
      meetingMode: [
        this.state.meetingMode,
        "Meeting Mode",
        "isWarningMeetingMode",
      ],
      meetingLink: [
        this.state.meetingLink,
        "Meeting Link",
        "isWarningMeetingLink",
      ],
      committeeMembers: [
        this.state.committeeMembersData,
        "Please select Committee Members",
        "isWarningCommitteeMembers",
      ],
      committeeGuestMembers: [
        this.state.committeeGuestMembersData,
        "Please select Guest Members",
        "isWarningCommitteeGuestMembers",
      ],
      committeeNoteRecords: [
        this.state.committeeNoteRecordsData,
        "Please select Committee Note Records",
        "isWarningCommitteeNoteRecords",
      ],
    };

    // console.log(fieldValues);
    const invalid: string[] = [];
    const warnings: Record<string, boolean> = {};

    // Check each field's value
    Object.values(fieldValues).forEach(
      ([value, displayName, warningKey]: any) => {
        // console.log(warningKey, value);

        if (
          value === null ||
          value === undefined ||
          value === "" ||
          (Array.isArray(value) && value.length === 0)
        ) {
          // console.error(`${displayName} (${warningKey}) is invalid.`);
          invalid.push(displayName as string); // Push invalid field name
          warnings[warningKey as string] = true; // Set warning to true for the invalid field
        } else {
          warnings[warningKey as string] = false; // Clear warning for valid fields
        }
      }
    );

    // Update state with invalid fields
    this.setState({ invalidFields: invalid, ...warnings });

    // Return false if there are any invalid fields
    const isValid = invalid.length === 0;
    // console.log(isValid);
    return isValid;
  };

  // import { sp } from "@pnp/sp";

  private getUserDepartmentByEmail = async (id: any): Promise<string | null> => {
    try {
      const userProfile = await this.props.sp.web.getUserById(id)();
      // console.log(userProfile);
  
      const profile = await this.props.sp.profiles.getPropertiesFor(userProfile?.LoginName);
      // console.log(profile.DisplayName);
      // console.log(profile.Email);
      // console.log(profile.Title);
  
      const departmentProperty = profile.UserProfileProperties.find(
        (element: any) => element.Key === "Department"
      );
  
      const department = departmentProperty?.Value || null;
      // console.log(department, "Department");
      this.setState({convernorFeildValue:department})
  
      return department;
    } catch (error) {
      console.error("Error fetching user profile:", error);
      return null;
    }
  };
  
// Example usage:
// getUserDepartmentByEmail("user@example.com");


  public render(): React.ReactElement<IXenWpCommitteeMeetingsFormsProps> {
    // console.log(this.props, "Props of Edit and Create Form while fetching");
    // console.log(this.state);

    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <div>
        {/* Title Seciton */}
        <div className={styles.titleContainer}>
          <div className={`${styles.noteTitle}
          
          `}>
            <div className={styles.statusContainer}>
              {this._itemId ? (
                <p className={styles.status}>
                  Status:
                  {
                    //   (():any=>{
                    //   switch(this.state.statusNumber){
                    //     case '1000':
                    //       console.log('Created')
                    //       return "Created"
                    //     case '2000':
                    //       return "Published"
                    //     case '3000':
                    //       return 'Meeting Over'
                    //     case '4000':
                    //       return 'Meeting Published'

                    //   }
                    // })()

                    this.state.MeetingStatus
                  }
                </p>
              ) : (
                ""
              )}
            </div>
            <h1 className={styles.title}>
              {this._itemId
                ? `eCommittee Meeting -${this.state.meetingId}`
                : `eCommittee Meeting -${this.props.formType}`}
            </h1>
            <p className={styles.titleDate}>
              {" "}
              <DateTime />
            </p>
          </div>

          <span className={styles.field}>
            All fields marked "*" are mandatory
          </span>
        </div>
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            General Section
          </h1>
        </div>

        <div
          className={`${styles.generalSection}`}
          style={{
            flexGrow: 1,
            margin: "10 10px",
            boxSizing: "border-box",
          }}
        >
          {/* Meeting ID: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting ID:
              <span className={styles.warning}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.meetingId}
            />
          </div>

          {/* Committee Name Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Committee Name :<span className={`${styles.warning}`}>*</span>
            </label>
            <Dropdown
              placeholder="Select an option"
              onRenderCaretDown={() =>
                this.onRenderCaretDowncommitteeNameFeildValue()
              }
              onChange={this.handleCommittename}
              className={styles.dropdown}
              options={this.state.committename}
              selectedKey={this.state.committeeNameFeildValue}
              styles={{
                dropdown: {
                  // width: 300,

                  border:
                    this.state.committeeNameFeildValue === "" &&
                    this.state.isWarningCommitteeName
                      ? "2px solid red"
                      : "1px solid transparent",
                },
              }}
            />
          </div>

          {/* Convenor Department : Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Convenor Department :<span className={styles.warning}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.convernorFeildValue}
              styles={{
                fieldGroup: {
                  border:
                    !this.state.convernorFeildValue &&
                    this.state.isWarningConvenor
                      ? "2px solid red"
                      : "1px solid",
                },
              }}
            />
          </div>

          {/* Chairman: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Chairman:
              <span className={styles.warning}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.charimanFeildValue}
              styles={{
                fieldGroup: {
                  border:
                    !this.state.charimanFeildValue &&
                    this.state.isWarningChairman
                      ? "2px solid red"
                      : "1px solid",
                },
              }}
            />
          </div>

          {/* Meeting Date: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Date :<span className={styles.warning}>*</span>
            </label>
            <DatePicker
              // firstDayOfWeek={firstDayOfWeek}
              placeholder="Select a date..."
              value={this.state.meetingDate}
              onSelectDate={this.onDateChange}
              ariaLabel="Select a date"
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
              styles={{
                root: {
                  height: "35px",
                  border:
                    this.state.meetingSubject === "" &&
                    this.state.isWarningMeetingSubject
                      ? "2px solid red"
                      : "",
                },
              }}
            />
          </div>

          {/* Meeting Subject: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Subject :<span className={styles.warning}>*</span>
            </label>
            <textarea
              className={styles.textarea}
              value={this.state.meetingSubject}
              onChange={(event) =>
                this.handleInputChange(event, "meetingSubject")
              }
              style={{
                border:
                  this.state.meetingSubject === "" &&
                  this.state.isWarningMeetingSubject
                    ? "2px solid red"
                    : "",
              }}
            />
          </div>

          {/* Meeting Mode : Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Mode :<span className={`${styles.warning}`}>*</span>
            </label>
            <Dropdown
              placeholder="Select an option"
              options={this.options}
              selectedKey={this.state.meetingMode}
              onChange={this.handleMeetingModeChange}
              className={styles.dropdown}
              styles={{
                dropdown: {
                  // width: 300,

                  border:
                    this.state.meetingMode === "" &&
                    this.state.isWarningMeetingMode
                      ? "2px solid red"
                      : "1px solid transparent",
                },
              }}
            />
          </div>

          {/* Meeting Link: Section */}
          <div className={styles.halfWidth} style={{paddingTop:'10px'}}>
            <label className={styles.label}>
              Meeting Link :<span className={styles.warning}>*</span>
            </label>
            <textarea
              className={styles.textareaForMeetingLink}
              value={this.state.meetingLink}
              onChange={(event) => this.handleInputChange(event, "meetingLink")}
              style={{
                border:
                  this.state.meetingLink === "" &&
                  this.state.isWarningMeetingLink
                    ? "2px solid red"
                    : "",
              }}
            />
          </div>
        </div>

        {/* Committee Members section */}

        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Members
          </h1>
        </div>
       
            <div
            className={`${styles.generalSectionApproverDetails}`}
            style={{ flexGrow: 1, margin: "10 10px" }}
          >
            <div>
              { (this.state.statusNumber=== '' || this.state.statusNumber=== '1000'|| this.state.statusNumber=== '2000')
            
            && 
            <div className={`${styles.peoplePickerAndSpanContainer}`}>
            <div style={{ display: "flex",flexWrap:'wrap' }}>
              <PeoplePicker
                key={this.state.committeeMemberskey}
                placeholder="Add Member..."
                context={this._peopplePicker}
                // titleText="People Picker"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                defaultSelectedUsers={[""]}
                disabled={false}
                ensureUser={true}
                onChange={this._getPeoplePickerItemsCommitteeMembers}
                // showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
              {/* <PeoplePicker /> */}
              <DefaultButton
                style={{ marginTop: "0px", marginLeft: "6px" }}
                type="button"
                className={`${styles.responsiveButton}`}
                onClick={(e) => {
                  if (this.state.selectedCommitteeMembers.length === 0) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage: "Please select Member then click on Add.",
                    });
                    this._clearCommitteeMembersPeoplePicker();
                    return;
                  }
                  this.handleOnAdd(e, "committeeMembers");
                }}
                iconProps={{ iconName: "Add" }}
              >
                Add
              </DefaultButton>
            </div>
            <span className={`${styles.spanForPeoplePicker}`}>
                  (Please enter minimum 3 character to search)
                </span>
          </div>}
              
            
              
            </div>
            <div style={{ overflowX: "auto" }}>
                <DetailsList
                  items={this.state.committeeMembersData} // Data for the table
                  columns={this.columnsCommitteeMembers} // Columns for the table
                  layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                  selectionMode={SelectionMode.none} // No selection column
                  isHeaderVisible={true} // Show column headers
                />
              </div>
          
          </div>

            

      
        {/* Committee Guest  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Guest Members
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            {(this.state.statusNumber=== '' || this.state.statusNumber=== '1000'|| this.state.statusNumber=== '2000')
            
            &&
            <div className={`${styles.peoplePickerAndSpanContainer}`}>
            <div style={{ display: "flex" ,flexWrap:'wrap'}}>
              <PeoplePicker
                key={this.state.committeeGuestMemberskey}
                placeholder="Add Member..."
                context={this._peopplePicker}
                // titleText="People Picker"
                personSelectionLimit={1}
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                defaultSelectedUsers={[""]}
                disabled={false}
                ensureUser={true}
                onChange={this._getPeoplePickerItemsCommitteeGuestMembers}
                // showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
              />
              {/* <PeoplePicker /> */}
              <DefaultButton
                style={{ marginTop: "0px", marginLeft: "6px" }}
                type="button"
                className={`${styles.responsiveButton}`}
                onClick={(e) => {
                  if (this.state.selectedCommitteeGuestMembers.length === 0) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage:
                        "Please select Guest member then click on Add.",
                    });
                    this._clearCommitteeMembersPeoplePicker();
                    return;
                  }

                  this.handleOnAdd(e, "committeeGuestMembers");
                }}
                iconProps={{ iconName: "Add" }}
              >
                Add
              </DefaultButton>
            </div>
            <span className={`${styles.spanForPeoplePicker}`}>
                  (Please enter minimum 3 character to search)
                </span>
          </div>}
           
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.committeeGuestMembersData} // Data for the table
                columns={this.columnsCommitteeGuestMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>
        {/* Committee Note  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          // style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Committee Note Records
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            {(this.state.statusNumber=== ''|| this.state.statusNumber=== '1000'|| this.state.statusNumber=== '2000')
            
            &&
            <div className={`${styles.peoplePickerAndSpanContainer}`}>
            <div style={{ display: "flex" ,flexWrap:'wrap'}}>
              <Dropdown
                selectedKey={this.state.committeeNoteRecordSelectedValue}
                onChange={this.handleCommitteeNoteRecordsDropdownChange}
                options={this.state.committeeNoteRecordDropDownData}
                placeholder="Add Note Record..."
                styles={{
                  root: {
                    minWidth: "180px",
                  },
                }}
              />
              {/* <PeoplePicker /> */}
              <DefaultButton
                style={{ marginTop: "0px", marginLeft: "6px" }}
                type="button"
                className={`${styles.responsiveButton}`}
                // onClick={(e) => this.handleOnAdd(e, "committeeNoteRecords")}
                iconProps={{ iconName: "Add" }}
                onClick={() => {
                  if (this.state.committeeNoteRecordSelectedValue === "") {
                    this.setState({
                      isModalOpen: true,
                      modalMessage: "Please select Note then click on Add.",
                    });
                    this._clearCommitteeMembersPeoplePicker();
                    return;
                  }
                  this._handleOnAddCommitteeNoteRecords();
                  this.setState({ committeeNoteRecordSelectedValue: "" });
                }}
              >
                Add
              </DefaultButton>
            </div>
           
          </div>}
          
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.committeeNoteRecordsData} // Data for the table
                columns={this.columnsCommitteeNoteRecords} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>


          {/* Comments section */}
          {this.state.statusNumber!== '' &&
             <div
             className={`${styles.generalSectionMainContainer}`}
             style={{ flexGrow: 1, margin: "10 10px" }}
           >
             <h1 className={styles.viewFormHeaderSectionContainer}>Comments</h1>
           </div>
          }
       

        {this.state.statusNumber!== '' && 
         <div
         className={`${styles.generalSectionApproverDetails}`}
         style={{ flexGrow: 1, margin: "10 10px" }}
       >
         <div>
           <div style={{ overflowX: "auto" }}>
             <DetailsList
               items={this.state.CommitteeMeetingMemberCommentsDT} // Data for the table
               columns={this.columnsCommitteeComments} // Columns for the table
               layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
               selectionMode={SelectionMode.none} // No selection column
               isHeaderVisible={true} // Show column headers
             />
           </div>
         </div>
       </div>}
       
        {/* WorkFlow  section */}

        {this._itemId ? (
          <>
            <div
              className={`${styles.generalSectionMainContainer}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <h1 className={styles.viewFormHeaderSectionContainer}>
                Workflow Log
              </h1>
            </div>
            <div
              className={`${styles.generalSectionApproverDetails}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <div>
                <div style={{ overflowX: "auto" }}>
                  <DetailsList
                    items={this.state.auditTrail} // Data for the table
                    columns={this.columnsCommitteeWorkFlowLog} // Columns for the table
                    layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                    selectionMode={SelectionMode.none} // No selection column
                    isHeaderVisible={true} // Show column headers
                  />
                </div>
              </div>
            </div>
          </>
        ) : (
          ""
        )}

        {/*  Buttons Section */}

        <div className={styles.buttonSectionContainer}>
          {this.state.statusNumber !== "1000" &&
            this.state.statusNumber !== "2000" &&
            this.state.statusNumber !== "3000" &&
            this.state.statusNumber !== "4000" &&
            this.state.statusNumber !== "5000" &&this.state.statusNumber !== "6000" && this.state.statusNumber !== "7000" &&this.state.statusNumber !== "9000" && (
              <PrimaryButton
                // type="button"
                onClick={() => {
                  if (!this._checkFields()) {
                    this.setState({
                      isModalOpen: true,
                      dialogType: "validation",
                      modalMessage: "Please Enter Reqired Fields",
                    });

                    return;
                  }
                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "createMeeting",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "PageHeaderEdit" }}
              >
                Create Meeting
              </PrimaryButton>
            )}

          {this.state.statusNumber !== "2000" &&
            this.state.statusNumber !== "3000" &&
            this.state.statusNumber !== "4000" &&
            this.state.statusNumber !== "5000"  &&this.state.statusNumber !== "6000" && this.state.statusNumber !== "7000" && this.state.statusNumber !== "9000" && (
              <PrimaryButton
                // type="button"

                onClick={() => {
                  if (!this._checkFields()) {
                    this.setState({
                      isModalOpen: true,
                      dialogType: "validation",
                      modalMessage: "Please Enter Reqired Fields",
                    });

                    return;
                  }

                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "publishMeeting",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "EntitlementRedemption" }}
              >
                Publish Meeting
              </PrimaryButton>
            )}


{this.state.statusNumber === '7000' && (
              <PrimaryButton
                // type="button"

                onClick={() => {
                  if (!this._checkFields()) {
                    this.setState({
                      isModalOpen: true,
                      dialogType: "validation",
                      modalMessage: "Please Enter Reqired Fields",
                    });

                    return;
                  }

                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "returnBack",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "EntitlementRedemption" }}
              >
                Return Back
              </PrimaryButton>
            )}

          {(this.state.statusNumber === "1000" ||
            this.state.statusNumber === "2000") && (
            <PrimaryButton
              // type="button"
              onClick={() => {
                if (this.state.statusNumber !== "2000") {
                  if (this._checkAllNoteRecord()) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage:
                        "Please publish the meeting to record meeting over.",
                    });

                    return;
                  }
                }
                this.setState({
                  isModalOpen: true,
                  dialogType: "confirmation",
                  btnType: "meetingOver",
                });
              }}
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "EntitlementRedemption" }}
            >
              Meeting Over
            </PrimaryButton>
          )}

          {this.state.statusNumber === "3000" &&
            this.state.statusNumber !== "4000" && (
              <PrimaryButton
                // type="button"
                onClick={() => {
                  if (this._checkAllNoteRecord()) {
                    this.setState({
                      isModalOpen: true,
                      modalMessage:
                        "MOM cannot be blank, Please add MOM for all the listed Notes.",
                    });

                    return;
                  }

                  this.setState({
                    isModalOpen: true,
                    dialogType: "confirmation",
                    btnType: "momPublished",
                  });
                }}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "EntitlementRedemption" }}
              >
                Mom Published
              </PrimaryButton>
            )}

          <DefaultButton
            // type="button"
            onClick={() => {
              const pageURL: string = this._homePageUrl;
              window.location.href = `${pageURL}`;
            }}
            className={`${styles.responsiveButton} `}
            iconProps={{ iconName: "Cancel" }}
          >
            Exit
          </DefaultButton>
        </div>

        {/* Modal for alerts */}
        <Modal
          isOpen={this.state.isModalOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={this.stylesModal.modal}
        >
          {this._getBodyContentOfDialogBox()}
        </Modal>

        <Modal
          isOpen={this.state.isModalMOMOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={this.stylesMOMModal.modal}
        >
          <>
            <div className={this.stylesMOMModal.header}>
              <div style={{ display: "flex", alignItems: "center" }}>
                <IconButton iconProps={{ iconName: "Info" }} />
                <h4 className={this.stylesMOMModal.headerTitle}>ADD MOM</h4>
              </div>
              <IconButton
                iconProps={{ iconName: "Cancel" }}
                onClick={() => {
                  this._closeModal();
                  this.setState({ dialogType: "" });
                }}
              />
            </div>
            <div 
            className={this.stylesMOMModal.body}
            >
              <div
              //  className={` ${styles.richTextContainer}`}
              >
                <RichText
                  value={this.state.draftResolutionFieldValue}
                  styleOptions={this.state.isSmallScreen?{showBold: true,
                    showItalic:true,showUnderline:true,showList:true,
                    showMore:true}:{
                      showBold: true,

                    showItalic:true,showUnderline:true,showList:true,
                    showAlign:true,
                    showImage:true,
                    showLink:true,
                    showStyles:true,

                    showMore:true

                    }}
                  onChange={(text: string) =>
                    this._onRichTextChangeForMom(text)
                  }

                  

                 
                />
              </div>
            </div>
            <div className={this.stylesMOMModal.footer}>
              <PrimaryButton
                iconProps={{ iconName: "Add" }}
                onClick={() => {
                  const updatedData = this.state.committeeNoteRecordsData.map(
                    (each: any) => {
                      if (each.key === this.state.selectedMOMNoteRecord) {
                        return {
                          ...each,
                          mom: this.state.draftResolutionFieldValue,
                        };
                      }
                      return each;
                    }
                  );

                  this.setState({
                    draftResolutionFieldValue: "",
                    committeeNoteRecordsData: updatedData,
                    dialogType: "",
                  });

                  this._closeModal();
                }}
                text="Add"
              />
            </div>
          </>
        </Modal>
        {/* Loading Section */}

        {this.state.isLoading && (
              <div>
                <Modal
                  isOpen={this.state.isLoading}
                  containerClassName={styles.spinnerModalTranparency}
                  styles={{
                    main: {
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      background: "transparent", // Removes background color
                      boxShadow: "none", // Removes box shadow
                    },
                  }}
                >
                  <div className="spinner">
                    <Spinner
                      label="still loading..."
                      ariaLive="assertive"
                      size={SpinnerSize.large}
                    />
                  </div>
                </Modal>
              </div>
            )}
      </div>
    );
  }
}
