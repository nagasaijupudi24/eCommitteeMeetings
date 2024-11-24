/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import styles from "./XenWpCommitteeMeetingsForms.module.scss";
import "./CustomStyles/custom.css";
import type { IXenWpCommitteeMeetingsFormsProps } from "./IXenWpCommitteeMeetingsFormsProps";
import {
  // DatePicker,
  DefaultButton,
  // defaultDatePickerStrings,
  DetailsList,
  DetailsListLayoutMode,
  Dialog,
  DialogFooter,
  DialogType,
  // Dropdown,
  IColumn,
  Icon,
  Link,
  PrimaryButton,
  SelectionMode,
  Stack,
  TextField,
  Toggle,
} from "@fluentui/react";
import { RichText } from "@pnp/spfx-controls-react/lib/controls/richText";
// import {
//   IPeoplePickerContext,
//   PeoplePicker,
//   PrincipalType,
// } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { escape } from '@microsoft/sp-lodash-subset';

interface CommtteeMeetingsState {
  MeetingNumber: string;
  MeetingDate: string;
  MeetingLink: string;
  MeetingMode: string;
  MeetingSubject: string;
  MeetingStatus: string;
  Department: string;
  ConsolidatedPDFPath: string;
  CommitteeName: string;
  Chairman: any;
  CommitteeMeetingGuestMembersDTO: any;
  CommitteeMeetingMembersDTO: any;
  CommitteeMeetingNoteDTO: any;
  CommitteeMeetingMembers: any;
  CommitteeMeetingGuests: any;
  AuditTrail: any;
  StatusNumber: string;
  CurrentApprover: any;
  FinalApprover: any;
  PreviousApprover: any;
  Confirmation: any;
  actionBtn: string;
  hideCnfirmationDialog: boolean;
  hideSuccussDialog: boolean;
  hideWarningDialog: boolean;
  SuccussMsg: string;
  CommitteeMeetingMemberCommentsDT: any;
  comments: string;
  isRturn: boolean;
  Created: any;
  departmentAlias: any;
  meetingId: any;
}
const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("itemId");
  return Number(Id);
};
const dragOptions = {
  moveMenuItemText: "Move",
  closeMenuItemText: "Close",
  // menu: ContextualMenu,
};
const modalPropsStyles = {
  main: {
    maxWidth: 600,
  },
};
const dialogContentProps = {
  type: DialogType.normal,
  title: "Alert",
  // subText: "Do you want to send this message without a subject?",
};
export default class XenWpCommitteeMeetingsViewForm extends React.Component<
  IXenWpCommitteeMeetingsFormsProps,
  CommtteeMeetingsState
> {
  private _listName;

  constructor(props: any) {
    super(props);
    this.state = {
      departmentAlias: "",
      meetingId: "",

      MeetingNumber: "",
      MeetingDate: "",
      MeetingLink: "",
      MeetingMode: "",
      MeetingSubject: "",
      MeetingStatus: "",
      Department: "",
      ConsolidatedPDFPath: "",
      CommitteeName: "",
      Chairman: null,
      CommitteeMeetingGuestMembersDTO: [],
      CommitteeMeetingMembersDTO: [],
      CommitteeMeetingNoteDTO: [],
      CommitteeMeetingMembers: [],
      CommitteeMeetingGuests: [],
      AuditTrail: [],
      StatusNumber: "",
      CurrentApprover: null,
      FinalApprover: null,
      PreviousApprover: null,
      Confirmation: {
        Confirmtext: "",
        Description: "",
      },
      actionBtn: "",
      hideCnfirmationDialog: true,
      hideSuccussDialog: true,
      hideWarningDialog: true,
      SuccussMsg: "",
      CommitteeMeetingMemberCommentsDT: [],
      comments: "",
      isRturn: false,
      Created: null,
    };
    const listName = this.props.listName;
    this._listName = listName?.title;
    // console.log(this._listName, this.props.listName, "onload");
    this._getItemBy();
    this._fetchDepartmentAlias();
  }

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

      // Step 2: Find the department entry where the Title or Department contains "Development"
      const specificDepartment = items.find(
        (each: any) =>
          each.Department.includes("Development") ||
          each.Title?.includes("Development")
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
          () => {
            // console.log(
            //   "Updated state with department alias:",
            //   this.state.departmentAlias
            // );
          }
        );
      } else {
        // console.log("No department found with 'Development' in title.");
      }
    } catch (error) {
      // console.error("Error fetching department alias: ", error);
    }
  };

  private _getItemBy = async () => {
    let user = await this.props.sp?.web.currentUser();
    // this._currentUser =user.id
    console.log(user, "user");
    const itemId = getIdFromUrl();
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(Number(itemId))
      .select(`*,Created,Author/Title,Author/EMail,
        Editor/Title,
        CurrentApprover/Title,
        CurrentApprover/EMail,
        CurrentApprover/JobTitle,
        FinalApprover/Title,
        FinalApprover/EMail,
        FinalApprover/JobTitle,
        PreviousApprover/Title,
        Chairman/Title,
        Chairman/EMail,
        PreviousApprover/EMail`).expand(`Author,Editor,
     CurrentApprover,PreviousApprover,FinalApprover,Chairman`)();
    console.log(item, "item");

    const currentyear = new Date().getFullYear();
    const nextYear = (currentyear + 1).toString().slice(-2);

    if (item) {
      // console.log(JSON.parse(item.AuditTrail));
      this.setState({
        meetingId: `${this.state.departmentAlias}/${currentyear}-${nextYear}/${itemId}`,
        MeetingNumber: item.MeetingNumber,
        MeetingDate: item.MeetingDate
          ? new Date(item.MeetingDate).toLocaleDateString()
          : "",
        MeetingLink: item.MeetingLink,
        MeetingMode: item.MeetingMode,
        MeetingSubject: item.MeetingSubject,
        MeetingStatus: item.MeetingStatus,
        Department: item.Department,
        ConsolidatedPDFPath: item.MeetingNumber,
        CommitteeName: item.CommitteeName,
        Chairman:
          item.Chairman === null && item.ChairmanId === null
            ? null
            : item.Chairman,
        CommitteeMeetingGuestMembersDTO:
          item.CommitteeMeetingGuestMembersDTO === null
            ? []
            : JSON.parse(item.CommitteeMeetingGuestMembersDTO),
        CommitteeMeetingMembersDTO:
          item.CommitteeMeetingMembersDTO === null
            ? []
            : JSON.parse(item.CommitteeMeetingMembersDTO), //CommitteeMeetingMemberCommentsDTO
        CommitteeMeetingMemberCommentsDT:
          item.CommitteeMeetingMemberCommentsDT === null
            ? []
            : JSON.parse(item.CommitteeMeetingMemberCommentsDT), //CommitteeMeetingMemberCommentsDTO
        CommitteeMeetingNoteDTO:
          item.CommitteeMeetingNoteDTO === null
            ? []
            : JSON.parse(item.CommitteeMeetingNoteDTO),
        CommitteeMeetingMembers:
          item.CommitteeMeetingMembers === null
            ? []
            : item.CommitteeMeetingGuestMembersDTO,
        CommitteeMeetingGuests: [],
        AuditTrail: item.AuditTrail === null ? [] : JSON.parse(item.AuditTrail),
        StatusNumber: item.StatusNumber,
        CurrentApprover:
          item.CurrentApprover === null && item.CurrentApproverId === null
            ? null
            : item.CurrentApprover,
        FinalApprover:
          item.FinalApprover === null && item.FinalApproverId === null
            ? null
            : item.FinalApprover,
        PreviousApprover:
          item.PreviousApprover === null && item.PreviousApproverId === null
            ? null
            : item.PreviousApprover,
        Created:
          new Date(item.Created).toLocaleDateString() +
          " " +
          new Date(item.Created).toLocaleTimeString(),
      });
    }
  };

  private columnsCommitteeMembers: IColumn[] = [
    {
      key: "memberName",
      name: "Member Name",
      fieldName: "memberEmailName",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "actionDate",
      name: "Action Date",
      fieldName: "actionDate",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
    },
  ];

  // private committeeMembersData = [
  //   {
  //     memberName: "John Doe",
  //     srNo: 1,
  //     designation: "Chairperson",
  //     actionDate: "2024-11-01",
  //   },
  //   {
  //     memberName: "Jane Smith",
  //     srNo: 2,
  //     designation: "Secretary",
  //     actionDate: "2024-11-05",
  //   },
  //   {
  //     memberName: "Michael Brown",
  //     srNo: 3,
  //     designation: "Treasurer",
  //     actionDate: "2024-11-10",
  //   },
  //   {
  //     memberName: "Emily Johnson",
  //     srNo: 4,
  //     designation: "Member",
  //     actionDate: "2024-11-15",
  //   },
  // ];

  private isReturnChecked = (
    event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    // console.log('toggle is ' + (checked ? 'checked' : 'not checked'));
    if (checked) {
      this.setState({
        isRturn: true,
      });
    } else {
      this.setState({ isRturn: false });
    }
  };

  private columnsCommitteeGuestMembers: IColumn[] = [
    {
      key: "guestMemberName",
      name: "Guest Members Name",
      fieldName: "memberEmailName",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "srNo",
      name: "SR No",
      fieldName: "srNo",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
    {
      key: "designation",
      name: "Designation",
      fieldName: "designation",
      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
    },
  ];

  // private committeeGuestMembersData = [
  //   {
  //     guestMemberName: "Alice White",
  //     srNo: 1,
  //     designation: "Advisor",
  //   },
  //   {
  //     guestMemberName: "Bob Green",
  //     srNo: 2,
  //     designation: "Consultant",
  //   },
  //   {
  //     guestMemberName: "Cathy Blue",
  //     srNo: 3,
  //     designation: "External Member",
  //   },
  //   {
  //     guestMemberName: "David Black",
  //     srNo: 4,
  //     designation: "Observer",
  //   },
  // ];

  private columnsCommitteeMeetingMinutes: IColumn[] = [
    {
      key: "serialNo",
      name: "S.No",
      fieldName: "serialNo",
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "noteTitle",
      name: "Note#",
      fieldName: "noteTitle",
      minWidth: 150,
      maxWidth: 200,
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
      key: "meetingMinutes",
      name: "Meeting Minutes",
      fieldName: "meetingMinutes",
      minWidth: 150,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: any) => (
        <RichText
          value={item.mom}
          isEditMode={false}
          style={{ minHeight: "auto", padding: "8px" }} // Adjusts height to content
        />
      ),
    },
    {
      key: "noteLink",
      name: "Note Link",
      fieldName: "noteLink",
      minWidth: 150,
      maxWidth: 250,
      isResizable: true,
      onRender(item, index, column) {
        return (
          <Link onClick={() => (window.location.href = `${item.noteLink}`)}>
            {item?.noteLink}
          </Link>
        );
      },
    },
  ];

  // private openDocuments=(link:any)=>{
  //   window.location.href=`${link}`

  // }

  // private committeeMeetingMinutesData = [
  //   {
  //     serialNo: 1,
  //     noteNumber: "001",
  //     committeeName: "Finance Committee",
  //     department: "Finance",
  //     meetingMinutes: "Discussed budget allocation",
  //     noteLink: "http://example.com/notes/001",
  //   },
  //   {
  //     serialNo: 2,
  //     noteNumber: "002",
  //     committeeName: "HR Committee",
  //     department: "Human Resources",
  //     meetingMinutes: "Reviewed new hiring policies",
  //     noteLink: "http://example.com/notes/002",
  //   },
  //   {
  //     serialNo: 3,
  //     noteNumber: "003",
  //     committeeName: "IT Committee",
  //     department: "Information Technology",
  //     meetingMinutes: "Discussed software upgrades",
  //     noteLink: "http://example.com/notes/003",
  //   },
  //   {
  //     serialNo: 4,
  //     noteNumber: "004",
  //     committeeName: "Marketing Committee",
  //     department: "Marketing",
  //     meetingMinutes: "Planned new campaign strategy",
  //     noteLink: "http://example.com/notes/004",
  //   },
  // ];

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

  // private committeeCommentsData = [
  //   {
  //     comments: "The project proposal is well-detailed and feasible.",
  //     commentedBy: "Alice White",
  //   },
  //   {
  //     comments: "Additional budget may be required for unexpected expenses.",
  //     commentedBy: "Bob Green",
  //   },
  //   {
  //     comments: "Consider involving external consultants for expert advice.",
  //     commentedBy: "Cathy Blue",
  //   },
  //   {
  //     comments: "Timeline seems tight; suggest extending by one month.",
  //     commentedBy: "David Black",
  //   },
  // ];

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

  // private getFormattedDate = (): string => {
  //   const { currentDate } = this.state;
  //   return `${currentDate.getDate()}-${
  //     currentDate.getMonth() + 1
  //   }-${currentDate.getFullYear()} ${currentDate.getHours()}:${currentDate.getMinutes()}:${currentDate.getSeconds()}`;
  // };

  private onClickMemberApprove = () => {
    this.setState({
      Confirmation: {
        Confirmtext: "Are you sure you want to approve this meeting?",
        Description: "Please click on Confirm button to approve meeting.",
      },
      hideCnfirmationDialog: false,

      actionBtn: "mbrApprove",
    });
  };
  private onClickMemberReturn = () => {
    if (this.state.comments) {
      this.setState({
        Confirmation: {
          Confirmtext: "Are you sure you want to retrun this meeting?",
          Description: "Please click on Confirm button to return meeting.",
        },
        hideCnfirmationDialog: false,
        actionBtn: "mbrReturn",
      });
    } else {
      this.setState({
        hideWarningDialog: false,
      });
    }
  };
  private onClickChairman = () => {
    this.setState({
      Confirmation: {
        Confirmtext: "Are you sure you want to approve this meeting?",
        Description: "Please click on Confirm button to approve meeting.",
      },
      hideCnfirmationDialog: false,
      actionBtn: "chairmanApprove",
    });
  };

  private handleApproveByMembers = async () => {
    const updatedCurrentApprover = this.state.CommitteeMeetingMembersDTO?.map(
      (obj: { memberEmail: any }) => {
        if (
          obj.memberEmail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase()
        ) {
          return {
            ...obj,
            status: "Approved",
            statusNumber: "9000",
            actionDate: new Date().toLocaleDateString(),
          };
        } else {
          return obj;
        }
      }
    );
    const isApprovedByAll = updatedCurrentApprover?.every(
      (obj: { status: string }) => obj.status === "Approved"
    );
    const auditTrail = this.state.AuditTrail;
    const comments = this.state.CommitteeMeetingMemberCommentsDT;

    auditTrail.push({
      action: `Committee meeting approved by ${this.props.userDisplayName}`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate:new Date().toLocaleDateString(),
    });
    console.log(comments)
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing: true,
        AuditTrail: JSON.stringify(auditTrail),
        // CommitteeMeetingMemberCommentsDT: this.state.comments
        //   ? JSON.stringify(comments)
        //   : null,
        MeetingStatus: isApprovedByAll
          ? "Pending Chairman Approval"
          : this.state.MeetingStatus,
        StatusNumber: isApprovedByAll ? "6000" : this.state.StatusNumber,
        CommitteeMeetingMembersDTO: JSON.stringify(updatedCurrentApprover),
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    if (item) {
      this.setState({
        hideSuccussDialog: !this.state.hideSuccussDialog,
        hideCnfirmationDialog: !this.state.hideCnfirmationDialog,
        SuccussMsg: "Committee meeting has been approved successfully",
      });
    }
  };
  private handleReturnByMembers = async () => {
    const updatedCurrentApprover = this.state.CommitteeMeetingMembersDTO?.map(
      (obj: { memberEmail: any }) => {
        if (
          obj.memberEmail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase()
        ) {
          return {
            ...obj,
            status: "Returned",
          };
        } else {
          return obj;
        }
      }
    );
    // const isApprovedByAll = updatedCurrentApprover?.every((obj: { status: string; })=>obj.status ==="Approved");
    const auditTrail = this.state.AuditTrail || [];
    const comments = this.state.CommitteeMeetingMemberCommentsDT;

    auditTrail.push({
      action: `Committee meeting returned by ${this.props.userDisplayName}`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate:new Date().toLocaleDateString(),
    });
    console.log(comments)
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing: true,
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingMemberCommentsDT: this.state.comments
          ? JSON.stringify(comments)
          : null,
        MeetingStatus: "Returned",
        StatusNumber: "7000",
        CommitteeMeetingMembersDTO: JSON.stringify(updatedCurrentApprover),
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    if (item) {
      this.setState({
        hideSuccussDialog: !this.state.hideSuccussDialog,
        hideCnfirmationDialog: !this.state.hideCnfirmationDialog,
        SuccussMsg: "Committee meeting has been returned successfully",
      });
    }
  };

  private handleApproveByChairman = async () => {
    const auditTrail: any[] = this.state.AuditTrail;
    const comments = this.state.CommitteeMeetingMemberCommentsDT;

    auditTrail.push({
      action: `Committee meeting approved by Chairman`,
      actionBy: this.props.userDisplayName,
      actionDate: new Date().toLocaleDateString(),
    });
    comments.push({
      comments: this.state.comments,
      commentedBy: this.props.userDisplayName,
      createdDate:new Date().toLocaleDateString(),
    });
    console.log(comments)
    const item = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.getById(getIdFromUrl())
      .update({
        startProcessing: true,
        AuditTrail: JSON.stringify(auditTrail),
        CommitteeMeetingMemberCommentsDT: this.state.comments
          ? JSON.stringify(comments)
          : null,

        // CommitteeMeetingMemberCommentsDTO: this.state.comments
        //   ? JSON.stringify(comments)
        //   : null,
        MeetingStatus: "Approved",
        StatusNumber: "9000",
        PreviousActionerId: (await this.props.sp?.web.currentUser())?.Id,
      });
    // Approved - 9000
    if (item) {
      this.setState({
        hideSuccussDialog: !this.state.hideSuccussDialog,
        hideCnfirmationDialog: !this.state.hideCnfirmationDialog,
        SuccussMsg: "Committee meeting has been approved successfully",
      });
    }
  };

  private handleComments = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: any
  ) => {
    this.setState({
      comments: newValue,
    });
  };

  private onConfirmation = () => {
    switch (this.state.actionBtn) {
      case "mbrApprove":
        this.handleApproveByMembers();
        break;
      case "mbrReturn":
        this.handleReturnByMembers();

        break;
      case "chairmanApprove":
        this.handleApproveByChairman();

        break;

      default:
        break;
    }
  };
  // private CreatedgetFormattedDate = (date: any): string => {
  //   const currentDate
  //   return `${currentDate.getDate()}-${
  //     currentDate.getMonth() + 1
  //   }-${currentDate.getFullYear()} ${currentDate.getHours()}:${currentDate.getMinutes()}:${currentDate.getSeconds()}`;
  // };

  public _checkCurrentApproverIsApprovedInCommitteMembersDTO = (): any => {
    const currentApprover = this.state.CommitteeMeetingMembersDTO.filter(
      (each: any) => {
        if (each.memberEmail === this.props.context.pageContext.user.email) {
          return each;
        }
      }
    );
    // console.log(currentApprover)
    return currentApprover[0]?.statusNumber !== "9000";
  };
  public render(): React.ReactElement<IXenWpCommitteeMeetingsFormsProps> {
    console.log(this.props,"prop ................")
    console.log(this.state)

    const modalProps: any = {
      isBlocking: true,
      styles: modalPropsStyles,
      dragOptions: dragOptions,
    };

    return (
      <div>
        {/* Title Seciton */}
        <div className={styles.titleContainer}>
          <div className={`${styles.noteTitleView} ${styles.commonProperties}`}>
            <div>
              {
                <p className={styles.status}>
                  Status: {this.state.MeetingStatus}{" "}
                </p>
              }
            </div>
            <h1 className={styles.title}>
              {getIdFromUrl()
                ? `eCommittee Meeting -${this.state.meetingId}`
                : `eCommittee Meeting -${this.props.formType}`}
            </h1>

            <p className={styles.titleDate}>Created : {this.state.Created}</p>
          </div>
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
              readOnly
            />
          </div>

          {/* Committee Name Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Committee Name :<span className={`${styles.warning}`}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.CommitteeName}
              readOnly
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
              value={this.state.Department}
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
              value={this.state.Chairman?.Title || ""}
            />
          </div>

          {/* Meeting Date: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Date :<span className={styles.warning}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.MeetingDate}
              readOnly
            />
            {/* <DatePicker
              // firstDayOfWeek={firstDayOfWeek}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              // DatePicker uses English strings by default. For localized apps, you must override this prop.
              strings={defaultDatePickerStrings}
            /> */}
          </div>

          {/* Meeting Subject: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Subject :<span className={styles.warning}>*</span>
            </label>
            <textarea
              className={styles.textarea}
              value={this.state.MeetingSubject}
              readOnly
            >
              {" "}
            </textarea>
          </div>

          {/* Meeting Mode : Sub Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Mode :<span className={`${styles.warning}`}>*</span>
            </label>
            <TextField
              type="text"
              className={styles.textField}
              value={this.state.MeetingMode}
              readOnly
            />
          </div>

          {/* Meeting Link: Section */}
          <div className={styles.halfWidth}>
            <label className={styles.label}>
              Meeting Link :<span className={styles.warning}>*</span>
            </label>
            <div className={styles.parentContainer}>
              <span
                className={styles.meetingLink}
                onClick={() => window.open(this.state.MeetingLink, "_blank")}
              >
                {this.state.MeetingLink}
              </span>
            </div>
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
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.CommitteeMeetingMembersDTO} // Data for the table
                columns={this.columnsCommitteeMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
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
            <div style={{ overflowX: "auto" }}>
              <DetailsList
                items={this.state.CommitteeMeetingGuestMembersDTO} // Data for the table
                columns={this.columnsCommitteeGuestMembers} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>
        {/* Meeting Minutes  section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>
            Meeting Minutes
          </h1>
        </div>
        <div
          className={`${styles.generalSectionApproverDetails}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <div>
            <div style={{ overflowX: "auto", width: "100%" }}>
              <DetailsList
                items={this.state.CommitteeMeetingNoteDTO} // Data for the table
                columns={this.columnsCommitteeMeetingMinutes} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>
        {this.state.CommitteeMeetingMembersDTO.some(
          (obj: { memberEmail: string }) =>
            obj.memberEmail.toLowerCase() ===
            this.props.context.pageContext.user.email.toLowerCase()
        ) &&
          this.state.StatusNumber === "5000" && (
            <div
              className={`${styles.generalSectionMainContainer}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <h1 className={styles.viewFormHeaderSectionContainer}>
                Comments section
              </h1>
            </div>
          )}
        {this.state.CommitteeMeetingMembersDTO.some(
          (obj: { memberEmail: string }) =>
            obj.memberEmail.toLowerCase() ===
            this.props.context.pageContext.user.email.toLowerCase()
        ) &&
          this.state.StatusNumber === "5000" && (
            <div
              className={`${styles.generalSectionApproverDetails}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <Toggle
                label="Do you want to return?"
                defaultChecked={false}
                onText="On"
                offText="Off"
                onChange={this.isReturnChecked}
                role="checkbox"
              />
              <br />
              {this.state.isRturn && (
                <div>
                   <label className={styles.label}>
                  Comments : 
                  
                 </label>
                <TextField
                  multiline
                  value={this.state.comments}
                  onChange={this.handleComments}
                  placeholder="Add Comment"
                ></TextField>

                </div>
                 
              )}
            </div>
          )}

        {this.state.Chairman?.EMail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase() &&
          this.state.StatusNumber === "6000" && (
            <div
              className={`${styles.generalSectionMainContainer}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
              <h1 className={styles.viewFormHeaderSectionContainer}>
                Comments section
              </h1>
            </div>
          )}
        {this.state.Chairman?.EMail.toLowerCase() ===
          this.props.context.pageContext.user.email.toLowerCase() &&
          this.state.StatusNumber === "6000" && (
            <div
              className={`${styles.generalSectionApproverDetails}`}
              style={{ flexGrow: 1, margin: "10 10px" }}
            >
               <label className={styles.label}>
             Comments : 
             
            </label>

              <TextField
                multiline
                value={this.state.comments}
                onChange={this.handleComments}
                placeholder="Add Comment"
              ></TextField>
            </div>
          )}

        {/* Comments section */}
        <div
          className={`${styles.generalSectionMainContainer}`}
          style={{ flexGrow: 1, margin: "10 10px" }}
        >
          <h1 className={styles.viewFormHeaderSectionContainer}>Comments</h1>
        </div>
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
        </div>

        {/* WorkFlow  section */}
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
                items={this.state.AuditTrail} // Data for the table
                columns={this.columnsCommitteeWorkFlowLog} // Columns for the table
                layoutMode={DetailsListLayoutMode.fixedColumns} // Keep columns fixed
                selectionMode={SelectionMode.none} // No selection column
                isHeaderVisible={true} // Show column headers
              />
            </div>
          </div>
        </div>

        {/*  Buttons Section */}

        <div className={styles.buttonSectionContainer}>
          {this._checkCurrentApproverIsApprovedInCommitteMembersDTO() && (
            <span
              hidden={
                !(
                  this.state.CommitteeMeetingMembersDTO.some(
                    (obj: { memberEmail: string }) =>
                      obj.memberEmail.toLowerCase() ===
                      this.props.context.pageContext.user.email.toLowerCase()
                  ) && this.state.StatusNumber === "5000"
                )
              }
            >
              <PrimaryButton
                onClick={this.onClickMemberApprove}
                className={`${styles.responsiveButton} `}
                iconProps={{ iconName: "DocumentApproval" }}
              >
                Approve
              </PrimaryButton>
            </span>
          )}

          <span
            hidden={
              !(
                this.state.CommitteeMeetingMembersDTO.some(
                  (obj: { memberEmail: string }) =>
                    obj.memberEmail.toLowerCase() ===
                    this.props.context.pageContext.user.email.toLowerCase()
                ) &&
                this.state.StatusNumber === "5000" &&
                this.state.isRturn
              )
            }
          >
            <PrimaryButton
              onClick={this.onClickMemberReturn}
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "ReturnToSession" }}
            >
              Return
            </PrimaryButton>
          </span>

          <span
            hidden={
              !(
                this.state.Chairman?.EMail.toLowerCase() ===
                  this.props.context.pageContext.user.email.toLowerCase() &&
                this.state.StatusNumber === "6000"
              )
            }
          >
            <PrimaryButton
              onClick={this.onClickChairman}
              className={`${styles.responsiveButton} `}
              iconProps={{ iconName: "DocumentApproval" }}
            >
              Approve
            </PrimaryButton>
          </span>

          <DefaultButton
            // type="button"
            onClick={() => {
              const pageURL: string = this.props.homePageUrl;
              window.location.href = `${pageURL}`;
            }}
            className={`${styles.responsiveButton} `}
            iconProps={{ iconName: "Cancel" }}
          >
            Exit
          </DefaultButton>
        </div>
        <Dialog
          hidden={this.state.hideCnfirmationDialog}
          onDismiss={() =>
            this.setState({
              hideCnfirmationDialog: !this.state.hideCnfirmationDialog,
            })
          }
          dialogContentProps={{
            ...dialogContentProps,
            title: (
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                style={{ padding: "0 20px" }}
              >
                <Icon iconName="Info" className={styles.dialogHeaderIcon} />
                <span style={{ fontSize: 18, fontWeight: "bold" }}>
                  Confirmation
                </span>
              </Stack>
            ),
          }}
          modalProps={modalProps}
          maxWidth={600}
        >
          {this.state.Confirmation && (
            <div className="dialogcontent_">
              <p>{this.state.Confirmation.Confirmtext}</p>
              <br />
              <p>{this.state.Confirmation.Description}</p>
            </div>
          )}

          <DialogFooter>
            <PrimaryButton
              onClick={this.onConfirmation}
              iconProps={{ iconName: "SkypeCircleCheck" }}
            >
              Confirm
            </PrimaryButton>
            <DefaultButton
              iconProps={{ iconName: "ErrorBadge" }}
              onClick={() =>
                this.setState({
                  hideCnfirmationDialog: !this.state.hideCnfirmationDialog,
                })
              }
            >
              Cancel
            </DefaultButton>
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.hideSuccussDialog}
          onDismiss={() =>
            this.setState({
              hideSuccussDialog: !this.state.hideSuccussDialog,
            })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
          maxWidth={600}
        >
          <div className="dialogcontent_">
            <p>{this.state.SuccussMsg}</p>
          </div>

          <DialogFooter>
            <PrimaryButton
              onClick={() => {
                const pageURL: string = this.props.homePageUrl;
                window.location.href = `${pageURL}`;
                this.setState({
                  hideSuccussDialog: !this.state.hideSuccussDialog,
                });
              }}
              iconProps={{ iconName: "ReturnToSession" }}
            >
              Ok
            </PrimaryButton>
          </DialogFooter>
        </Dialog>
        <Dialog
          hidden={this.state.hideWarningDialog}
          onDismiss={() =>
            this.setState({
              hideWarningDialog: !this.state.hideWarningDialog,
            })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
          maxWidth={600}
        >
          <div className="dialogcontent_">
            <p>Please fill in comments then click on return</p>
          </div>

          <DialogFooter>
            <PrimaryButton
              onClick={() =>
                this.setState({
                  hideWarningDialog: !this.state.hideWarningDialog,
                })
              }
              iconProps={{ iconName: "ReturnToSession" }}
            >
              Ok
            </PrimaryButton>
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
}
