import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import "bootstrap/dist/css/bootstrap.min.css";
import logo from "../assets/sona-comstarlogo.png";

interface IProps {
  context: any;
  itemId: number;
}

interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const ApproverAdvanceForm: React.FC<IProps> = ({ context, itemId }) => {
  const sp = spfi().using(SPFx(context));
  const [selectedVendorCode, setSelectedVendorCode] = useState("");
const actionLock = React.useRef(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [employee, setEmployee] = React.useState<any>({});
  const [selectedVendorName, setSelectedVendorName] = useState("");
  //const [itemData, setItemData] = useState<any>(null);
  const [approverRemarks, setApproverRemarks] = useState("");
  const [attachments, setAttachments] = useState<any[]>([]);
  //const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [itemData, setItemData] = useState<any>(null);
  //const [itemData, setItemData] = useState<any>(formData);
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");
  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
      if (!vendorId) {
        setPreviousAdvances([]);
        return;
      }
      console.log("Fetching for Vendor:", vendorId);

      const data = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.select(
          "PONumber",
          "RequestAdvanceAmount",
          "Created",
          "VoucherDate",

          "PaidAmount",
          "Status",
          "VendorCode/Id",
        )
        .expand("VendorCode")
        .filter(`VendorCode/Id eq ${vendorId} and Status eq 'Paid'`)
        .orderBy("Created", false)();

      console.log("DATA:", data);

      void setPreviousAdvances(data);
    } catch (error) {
      console.error("Error fetching previous advances:", error);
      void setPreviousAdvances([]);
    }
  };

  const buildApprovalPreview = async (employee: any) => {
    const flow: any[] = [];

    // RM
    if (employee.ReportingManager?.Title) {
      flow.push({
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Status: "Pending",
      });
    }

    // HOD
    if (employee.HOD?.Title) {
      flow.push({
        Name: employee.HOD.Title,
        Role: "HOD",
        Status: "Pending",
      });
    }

    // 🔥 Matrix from list
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Title")
      .expand("Approver,Role")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any) => ({
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Status: "Pending",
    }));

    setApprovalMatrix([...flow, ...matrixApprovers]);
  };

  const getLoggedInUser = async () => {
    try {
      debugger;
      const currentUser = await sp.web.currentUser();
      const email = currentUser.Email;

      const user = await sp.web.lists
        .getByTitle("EmployeeMaster")
        .items.select(
          "EmployeeCode",
          "EmployeeName",
          "Division",
          "Location",
          "EmployeeEmail",
          "ReportingManager/Title",
          "ReportingManager/Id",
          "HOD/Title",
          "HOD/Id",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        void setEmployee(user[0]);
      }
      buildApprovalPreview(user[0]);
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  const getAttachments = async (capexId: string) => {
    try {
      debugger;
      if (!capexId) return;

      console.log("CapexId:", capexId);

      const safeCapexId = capexId.replace(/\//g, "_");

      const folderPath = `CapexPaymentDocs/${safeCapexId}`;

      console.log("Folder Path:", folderPath);

      const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files();

      console.log("Files:", files);

      setAttachments(files || []);
    } catch (error) {
      console.log("Attachment fetch error:", error);
      setAttachments([]);
    }
  };

  const getVendors = async () => {
    try {
      const data = await sp.web.lists
        .getByTitle("VendorMaster")
        .items.select("Id", "VendorCode", "VendorName")();

      setVendors(data);
    } catch (error) {
      console.error("Vendor fetch error:", error);
    }
  };

  // ✅ Fetch Item by ID

  // useEffect(() => {
  //   debugger;
  //   if (!formData) return;
  //   debugger;
  //   console.log(formData);
  //   setSelectedVendorId(Number(formData.VendorCode));

  //   setItemData(formData);
  //   setApproverRemarks(formData.ApproverRemarks || "");

  //   // ✅ CALL ATTACHMENT FUNCTION HERE
  //   if (formData.CapexId) {
  //     getAttachments(formData.CapexId);
  //   }

  //   // Approval Matrix
  //   if (formData.ApprovalMatrix) {
  //     try {
  //       setApprovalMatrix(JSON.parse(formData.ApprovalMatrix));
  //     } catch {
  //       setApprovalMatrix([]);
  //     }
  //   }

  //   // Workflow History
  //   // Workflow History
  //   if (formData.WorkflowHistory) {
  //     try {

  //       const history =
  //         typeof formData.WorkflowHistory === "string"
  //           ? JSON.parse(formData.WorkflowHistory)
  //           : formData.WorkflowHistory || [];

  //       setWorkflowHistory(history);

  //       console.log("🔥 Workflow History:", history);

  //     } catch (e) {

  //       console.log("WFHistory parse error", e);

  //       setWorkflowHistory([]);
  //     }
  //   }
  // }, [formData]);
  const getItemById = async (itemId: any) => {
    try {
      debugger;
      console.log("Fetching item with ID:", itemId);

      const item = await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemId)
        .select(
          "ID",
          "Title",
          "Created",
          "EmployeeName",
          "Email",
          "EmployeeCode",
          "ContactNo",
          "EmployeeStatus",
          "Division",
          "Location",
          "RM",
          "HOD",
          "VendorCode/Id",
          "VendorCode/Title",
          "VendorName",
          "PODate",
          "POPaymentTerms",
          "POAmount",
          "MRNNumber",
          "MRNDtae",
          "MRNAmountwithGST",
          "PONumber",
          "RequestedAmountforPayment",
          "Status",
          "CurrentApproverId",
          "CapexId",
          "InstallationDetails",
          "FinalPaymentAgainstPO",
          "ApprovalMatrix",
          "WorkflowHistory",
          
         "RequestorName"

        )();

      console.log("ITEM DATA:", item.RequestorName);
      debugger;
      setItemData(item);
      debugger;

      const vendorId = item?.VendorCode?.Id || null;

      console.log("Vendor Id:", vendorId);

      setSelectedVendorId(vendorId);

      setSelectedVendorName(item?.VendorName || "");

      if (item.CapexId) {
        await getAttachments(item.CapexId);
      }

      if (item.ApprovalMatrix) {
        try {
          const parsed =
            typeof item.ApprovalMatrix === "string"
              ? JSON.parse(item.ApprovalMatrix)
              : item.ApprovalMatrix;

          setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("ApprovalMatrix parse error", e);
          setApprovalMatrix([]);
        }
      } else {
        setApprovalMatrix([]);
      }

      if (item.WorkflowHistory) {
        try {
          const parsed =
            typeof item.WorkflowHistory === "string"
              ? JSON.parse(item.WorkflowHistory)
              : item.WorkflowHistory;

          setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
        } catch (e) {
          console.error("WorkFlowHistory parse error", e);
          setWorkflowHistory([]);
        }
      } else {
        setWorkflowHistory([]);
      }
    } catch (error) {
      console.error("Fetch error:", error);
    }
  };

  useEffect(() => {
    debugger;
    if (!context || !itemId) return;

    const loadData = async () => {
      debugger;

     // await getLoggedInUser();

      await getVendors();

      await getItemById(itemId);
    };

    void loadData();
  }, [context, itemId]);

  useEffect(() => {
    if (selectedVendorId) {
      console.log("Calling Previous Advances:", selectedVendorId);

      void getPreviousAdvances(selectedVendorId);
    }
  }, [selectedVendorId]);

  const handleApprove = async () => {
     if (actionLock.current) return;

    actionLock.current = true;

    debugger;
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
        // actionLock.current = false;

        return;
      }

      // if (isProcessing) return;

      //setIsProcessing(true);

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentUser = context.pageContext.user.displayName;

      const currentIndex = flow.findIndex((a: any) => a.Name === currentUser);
      if (currentIndex === -1) {
        alert("You are not current approver");
        //// setIsProcessing(false);
        //actionLock.current = false;
        return;
      }

      flow[currentIndex].Status = "Approved";

      let nextApproverId = null;

      if (flow[currentIndex + 1]) {
        flow[currentIndex + 1].Status = "In Progress";
        nextApproverId = flow[currentIndex + 1].Id;
      }

      let finalStatus = itemData.Status;

      const currentRole = flow[currentIndex]?.Role;

      if (currentRole === "RM") {
        finalStatus = "Pending for Approver";
      } else if (currentRole === "HOD") {
        finalStatus = "Pending for PF Approver";
      } else {
        finalStatus = nextApproverId ? "Pending" : "Approved";
      }

      // let ApproverStatus;

      // if (currentRole === "RM") {
      //   ApproverStatus = "Pending for HOD";
      // } else if (currentRole === "HOD") {
      //   ApproverStatus = "Pending for PF Approver";
      // } else {
      //   ApproverStatus = nextApproverId ? "Pending" : "Approved";
      // }

      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Approved",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,
          Status: finalStatus,
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: nextApproverId,
          WorkflowHistory: JSON.stringify(history),
         // ApproverStatus: ApproverStatus,
        });

      alert("Approved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Approver";
    } catch (error) {
      console.error(error);
      alert("Error ❌");
      ////actionLock.current = false;

      //setIsProcessing(false);
    }
    finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };

  const handleApprove1 = async () => {
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
        // actionLock.current = false;

        return;
      }

      // =========================
      // 🔥 PARSE MATRIX
      // =========================
      let flow: any[] = [];

      try {
        flow =
          typeof itemData.ApprovalMatrix === "string"
            ? JSON.parse(itemData.ApprovalMatrix)
            : itemData.ApprovalMatrix || [];
      } catch {
        flow = [];
      }

      if (!Array.isArray(flow) || flow.length === 0) {
        alert("Approval Matrix empty");
        return;
      }

      // =========================
      // 🔥 FIND CURRENT APPROVER
      // =========================
      // const currentIndex = flow.findIndex(
      //   (x: any) => x.Status === "In Progress" || x.Status === "Pending",
      // );

      const currentUserId = context.pageContext.legacyPageContext.userId;

    const currentIndex = flow.findIndex(
      (a: any) => Number(a.Id) === Number(currentUserId)
    );


      if (currentIndex === -1) {
        alert("No pending approver found");
        return;
      }

      // =========================
      // 🔥 CURRENT APPROVED
      // =========================
      flow[currentIndex].Status = "Approved";

      // =========================
      // 🔥 NEXT APPROVER
      // =========================

      const nextIndex = currentIndex + 1;

      let nextApproverId: number | null = null;
      let nextStatus = "Approved";
      let pendingAt = "Completed";

      if (nextIndex < flow.length) {
        flow[nextIndex].Status = "Pending";

        nextApproverId = flow[nextIndex].Id;

        // =========================
        // 🔥 DYNAMIC STATUS
        // =========================
        const nextRole = flow[nextIndex].Role;

        if (nextRole === "RM") {
          nextStatus = "Pending for Approval";
        } else if (nextRole === "HOD") {
          nextStatus = "Pending for Approval";
        } else if (nextRole === "Performer") {
          nextStatus = "Pending for PF Approver";
        }

        pendingAt = `Pending at ${flow[nextIndex].Role}`;
      }

      // =========================
      // 🔥 WORKFLOW HISTORY
      // =========================
      let history: any[] = [];

      try {
        history =
          typeof itemData.WorkflowHistory === "string"
            ? JSON.parse(itemData.WorkflowHistory)
            : itemData.WorkflowHistory || [];
      } catch {
        history = [];
      }

      history.push({
        CurrentApprover: context.pageContext.user.displayName,

        ActionTaken: "Approved",

        Comment: approverRemarks,

        Date: new Date().toISOString(),

        CurrentStatus: pendingAt,
      });

      console.log("🔥 UPDATED MATRIX:", flow);
      console.log("🔥 UPDATED HISTORY:", history);

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,

          Status: nextStatus,

          PendingAt: pendingAt,

          CurrentApproverId: nextApproverId,

          ApprovalMatrix: JSON.stringify(flow),

          WorkflowHistory: JSON.stringify(history),
        });

      alert("Approved successfully ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Approver";
    } catch (error) {
      console.error("Approve error:", error);

      alert("Error ❌");
    }
  };

  // ✅ Sent Back
  const handleSendBack = async () => {
     if (actionLock.current) return;
   // actionLock.current = true;
    setIsProcessing(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
        //actionLock.current = false;
setIsProcessing(false);
        return;
      }
      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      //  const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);
      const currentUser = context.pageContext.user.displayName;

      const currentIndex = flow.findIndex((a: any) => a.Name === currentUser);

      if (currentIndex === -1) {
        alert("You are not current approver");
        // setIsProcessing(false);
        return;
      }

      flow[currentIndex].Status = "Send Back";

      let previousApproverId = null;

      if (flow[currentIndex - 1]) {
        flow[currentIndex - 1].Status = "In Progress";
        previousApproverId = flow[currentIndex - 1].Id;
      }

      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Send Back",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });

      // =========================
      // 🔥 UPDATE SHAREPOINT
      // =========================
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,

          Status: "Send Back",

          ApprovalMatrix: JSON.stringify(flow),

          CurrentApproverId: previousApproverId,

          WorkflowHistory: JSON.stringify(history),
        });

      alert("Send Back ✅");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Approver";
    } catch (error) {
      console.error(error);
      alert("Error ❌");
    }
     finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };

  // ✅ Reject
  const handleReject = async () => {
     if (actionLock.current) return;
   // actionLock.current = true;
    setIsProcessing(true);
    try {
      if (!approverRemarks || approverRemarks.trim() === "") {
        alert("Please enter Remarks");
        return;
      }

      const flow = itemData.ApprovalMatrix
        ? JSON.parse(itemData.ApprovalMatrix)
        : [];

      const currentUserId = context.pageContext.legacyPageContext.userId;

      const currentUser = context.pageContext.user.displayName;

      const currentIndex = flow.findIndex((a: any) => a.Name === currentUser);
      //  const currentIndex = flow.findIndex((a: any) => a.Id === currentUserId);

      if (currentIndex === -1) {
        alert("You are not current approver");
        //setIsProcessing(false);
        return;
      }

      flow[currentIndex].Status = "Rejected";

      const history = itemData.WorkflowHistory
        ? JSON.parse(itemData.WorkflowHistory)
        : [];

      history.push({
        CurrentApprover: context.pageContext.user.displayName,
        ActionTaken: "Rejected",
        Comment: approverRemarks,
        Date: new Date().toISOString(),
      });
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(itemData.ID)
        .update({
          ApproverRemarks: approverRemarks,

          Status: "Rejected",

          ApprovalMatrix: JSON.stringify(flow),

          CurrentApproverId: null, // 🔥 stop flow

          WorkflowHistory: JSON.stringify(history),
        });

      alert("Rejected ❌");

      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Approver";
    } catch (error) {
      console.error(error);
      alert("Error ❌");
    }
     finally {
    //actionLock.current = false;
      setIsProcessing(false);
  }
  };
  const handleExit = () => {
    window.location.href = `https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=Approver`;
  };

  // ⛔ Wait until data loads
  if (!itemData) return <div>Loading...</div>;

  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            <div className="bordered">
              <img src={logo} />
              <h1>Capex Payment Approval form </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  <li className={`approval-step`}>
                    {`Initiator`} - {itemData?.EmployeeName}
                  </li>
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${
                        a.Status === "In Progress"
                          ? "active"
                          : a.Status === "Approved"
                            ? "approved"
                            : a.Status === "Rejected"
                              ? "rejected"
                              : ""
                      }`}
                    >
                      {a.Role} - {a.Name}
                    </li>
                  ))}
                </ul>
              </div>
            )}
            <div className="borderedbox">
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Email}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {itemData.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.RM}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {itemData.HOD}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font"> Vendor Code </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.VendorCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.VendorName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.PONumber}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date </label> : &nbsp;&nbsp;
                    <label className="fonttext ">
                      {" "}
                      {itemData.PODate
                        ? new Date(itemData.PODate).toLocaleDateString("en-GB")
                        : ""}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Terms </label> : &nbsp;&nbsp;
                    <label className="fonttext ">
                      {" "}
                      {itemData.POPaymentTerms}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount </label> : &nbsp;&nbsp;
                    <label className="fonttext "> {itemData.POAmount}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>MRN & Payment Details </label> : &nbsp;&nbsp;
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number </label> : &nbsp;&nbsp;
                    <label className="fonttext">{itemData?.MRNNumber}</label>
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNDtae
                        ? new Date(itemData.MRNDtae).toLocaleDateString("en-GB")
                        : ""}
                    </label>
                    {/* <label className="fonttext">{itemData?.mrnDate}</label> */}
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.MRNAmountwithGST}
                    </label>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">
                      {itemData?.RequestedAmountforPayment}
                    </label>
                  </div>
                </div>
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="heading1" style={{ marginTop: "10px" }}>
                    <label>Final Payment Details</label>
                  </div>

                  <div className="main-formcontainer">
                    <div className="row mb-20">
                      <div className="col-md-4">
                        <label className="font">Final Payment Against PO</label>{" "}
                        : &nbsp;&nbsp;
                        <label className="fonttext">
                          {itemData?.FinalPaymentAgainstPO ? "Yes" : "No"}
                        </label>
                      </div>
                    </div>

                    {itemData?.FinalPaymentAgainstPO && (
                      <div className="row mb-20">
                        <div className="col-md-6">
                          <label className="font">Installation Details</label> :
                          &nbsp;&nbsp;
                          <label className="fonttext">
                            {itemData?.InstallationDetails}
                          </label>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Workflow History</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    {workflowHistory.length === 0 ? (
                      <p>No history available</p>
                    ) : (
                      <div className="workflow-history">
                        <table
                          className="workflow-table"
                          style={{ width: "100%" }}
                        >
                          <thead>
                            <tr>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action By
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Action Taken
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Date
                              </th>
                              <th style={{ padding: "8px", textAlign: "left" }}>
                                Comment
                              </th>
                            </tr>
                          </thead>
                          <tbody>
                            {workflowHistory
                              .filter(
                                (h: any) =>
                                  h.ActionTaken &&
                                  h.ActionTaken !== "Draft Saved" &&
                                  h.ActionTaken !== "Edited",
                              )
                              .map((h: any, idx: number) => (
                                <tr key={idx}>
                                  <td style={{ padding: "8px" }}>
                                    {h.CurrentApprover || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.ActionTaken || ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Date
                                      ? new Date(h.Date).toLocaleDateString(
                                          "en-GB",
                                        )
                                      : ""}
                                  </td>

                                  <td style={{ padding: "8px" }}>
                                    {h.Comment || ""}
                                  </td>
                                </tr>
                              ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Approver Action</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Approver Remarks</label>
                    <textarea
                      className="font-control"
                      value={approverRemarks}
                      onChange={(e) => setApproverRemarks(e.target.value)}
                    />
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Document</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments</label>
                    {attachments.length === 0 ? (
                      <p>No attachments found</p>
                    ) : (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a href={file.ServerRelativeUrl} target="_blank">
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
                    )}
                  </div>
                </div>
              </div>
              <div className="row my-3">
                <div className="col-md-12">
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "5px",
                    }}
                  >
                    <a className="submit-btn" onClick={handleApprove}>
                      Approve
                    </a>

                    <a className="Rework-btn" onClick={handleSendBack}>
                      Send Back
                    </a>

                    <a className="Reject-btn" onClick={handleReject}>
                      Reject
                    </a>

                    <a href="#" onClick={handleExit} className="reset-btn">
                      Exit
                    </a>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ApproverAdvanceForm;
