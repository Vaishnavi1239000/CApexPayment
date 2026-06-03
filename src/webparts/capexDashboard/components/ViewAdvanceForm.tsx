import * as React from "react";
import "./advanced.scss";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { useEffect, useState } from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import 'bootstrap/dist/css/bootstrap.min.css';
import logo from "../assets/sona-comstarlogo.png";
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}
const ViewAdvanceForm = ({ context, formData, onClose }: any) => {
  const [attachments, setAttachments] = useState<any[]>([]);
  const sp = spfi().using(SPFx(context));
  const [employee, setEmployee] = useState<any>({});
  // 🔹 Employee
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");

  const [finalPayment, setFinalPayment] = useState("");
  const [installationDetails, setInstallationDetails] = useState("");

  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [vendors, setVendors] = useState<IVendor[]>([]);
  //const [employeeName, setEmployeeName] = useState("");
  ////const [employeeEmail, setEmployeeEmail] = useState("");
  //const [VendorCode, setVendorCode] = useState<number | null>(null);
  // 🔹 Form fields
  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");

  const [vendorName, setVendorName] = useState("");
  const [glCode, setGlCode] = useState("");
  const [costCenter, setCostCenter] = useState("");
  const [remarks, setRemarks] = useState("");
  const [projectDesc, setProjectDesc] = useState("");

  // 🔹 Extra fields
  const [approverRemarks, setApproverRemarks] = useState("");
  const [voucherDate, setVoucherDate] = useState("");
  const [VouchingNumber, setVouchingNumber] = useState("");
  const [UTRDate, setUTRDate] = useState("");
  const [UTRNumber, setUTRNumber] = useState("");
  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };
  const getAttachments = async (capexId: string) => {
    try {
      if (!capexId) return;

      // ✅ IMPORTANT FIX
      const safeCapexId = capexId.replace(/\//g, "_");

      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safeCapexId}`;

      console.log("Correct Folder Path:", folderPath);

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

  // const uploadFiles = async () => {
  //     if (!formData?.CapexID || selectedFiles.length === 0) return;

  //     const safe = formData.CapexID.replace(/\//g, "_");
  //     const path = `/sites/SonaFinance/CapexAdvanceDocs/${safe}`;

  //     for (const file of selectedFiles) {
  //       await sp.web
  //         .getFolderByServerRelativePath(path)
  //         .files.addUsingPath(file.name, file, { Overwrite: true });
  //     }

  //   void  setSelectedFiles([]);
  //   void  getAttachments(formData.CapexID);
  //   };

  // ✅ Fetch Item by ID
  const getVendors = async () => {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName")();
    void setVendors(data);
  };

  // ✅ Bind SharePoint Data
  useEffect(() => {
    debugger;
    console.log(formData);
    if (!formData) return;

    setPoNumber(formData.PONumber || "");
    setPoDate(formData.PODate?.split("T")[0] || "");
    setPoTerms(formData.POPaymentTerms || "");
    setPoAmount(formData.POAmount || "");

    setVendorName(formData.VendorName || "");
    setSelectedVendorId(formData.VendorCode || null); // ✅ ADD THIS
    setSelectedVendorName(formData.VendorName || ""); // ✅ ADD THIS
    setMrnNumber(formData.MRNNumber || "");
    setMrnDate(formData.MRNDtae?.split("T")[0] || "");
    setMrnAmount(formData.MRNAmountwithGST || "");
    setRequestedAmount(formData.RequestedAmountforPayment || "");

    // ✅ Boolean → Yes/No
    setFinalPayment(formData.FinalPaymentAgainstPO ? "Yes" : "No");

    setInstallationDetails(formData.InstallationDetails || "");

    setApproverRemarks(formData.ApproverRemarks || "");
    setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    setVouchingNumber(formData.VouchingNumber || "");
    setUTRDate(formData.UTRDate?.split("T")[0] || "");
    setUTRNumber(formData.UTRNumber || "");

    // ✅ PIC FIX
    if (formData?.PICName?.Title) {
      setSelectedUser([
        {
          text: formData.PICName.Title,
          secondaryText: formData.PICName.EMail,
        },
      ]);
    }

    if (formData.CapexId) {
      void getAttachments(formData.CapexId);
    }
    if (formData?.ApprovalMatrix) {
      try {
        const parsed =
          typeof formData.ApprovalMatrix === "string"
            ? JSON.parse(formData.ApprovalMatrix)
            : formData.ApprovalMatrix;

        setApprovalMatrix(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("ApprovalMatrix parse error", e);
        setApprovalMatrix([]);
      }
    } else {
      setApprovalMatrix([]);
    }
    // ✅ Workflow History
    if (formData?.WorkFlowHistory) {
      try {
        const parsed =
          typeof formData.WorkFlowHistory === "string"
            ? JSON.parse(formData.WorkFlowHistory)
            : formData.WorkFlowHistory;

        setWorkflowHistory(Array.isArray(parsed) ? parsed : []);
      } catch (e) {
        console.error("WorkFlowHistory parse error", e);
        setWorkflowHistory([]);
      }
    } else {
      setWorkflowHistory([]);
    }
  }, [formData]);

  const handleExit = () => {
    if (onClose) onClose();
    else window.location.reload();
  };
  const getLoggedInUser = async () => {
    try {
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
          "HOD/Title",
          "ContactNo",
          "EmployeeStatus",
          "CostCenter",
        )
        .expand("ReportingManager", "HOD")
        .filter(`EmployeeEmail eq '${email}'`)
        .top(1)();

      if (user.length > 0) {
        setEmployee(user[0]);
      }
    } catch (error) {
      console.log("Error fetching user:", error);
    }
  };
  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
  }, []);
  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1> Advance Payment (View) </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <></>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${a.Status === "In Progress"
                        ? "active"
                        : a.Status === "Approved"
                          ? "approved"
                          : a.Status === "Rejected"
                            ? "rejected"
                            : a.Status === "Send Back"
                              ? "sendback"
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
              <div className="heading1">
                <label>Requestor Information</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Employee Code" className="font">
                      Employee Code
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeCode}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Name" className="font">
                      Employee Name{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.EmployeeName}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Email" className="font">
                      Employee Email{" "}
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeEmail}
                    </label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Contact No" className="font">
                      Contact No
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.ContactNo}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Employee Status" className="font">
                      Employee Status
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.EmployeeStatus}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="Division" className="font">
                      Division
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Division}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label htmlFor="Location" className="font">
                      Location
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.Location}</label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="RM" className="font">
                      RM
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">
                      {" "}
                      {employee.ReportingManager?.Title}
                    </label>
                  </div>
                  <div className="col-md-4">
                    <label htmlFor="HOD" className="font">
                      HOD
                    </label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext"> {employee.HOD?.Title}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Vendor & PO</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    {/* <select
                      value={selectedVendorId ?? ""}
                      disabled={true}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select> */}

                    <label className="fonttext">{selectedVendorId}</label>

                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label> : &nbsp;&nbsp;
                    <label className="fonttext">{vendorName}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label> : &nbsp;&nbsp;
                    <label className="fonttext">{poNumber}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">{poDate}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{poAmount}</label>
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Payment Terms</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{poTerms}</label>
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>MRN & Payment Details</label>
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number</label> : &nbsp;&nbsp;
                    <label className="fonttext">{mrnNumber}</label>
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">{mrnDate}</label>
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Amount</label> : &nbsp;&nbsp;
                    <label className="fonttext">{mrnAmount}</label>
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{requestedAmount}</label>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Final Payment Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Final Payment Against PO</label>{" "}
                    : &nbsp;&nbsp;
                    <label className="fonttext">{finalPayment}</label>
                  </div>
                </div>

                {finalPayment === "Yes" && (
                  <div className="row mb-20">
                    <div className="col-md-6">
                      <label className="font">Installation Details</label> :
                      &nbsp;&nbsp;
                      <label className="fonttext">
                        {installationDetails}
                      </label>
                    </div>
                  </div>
                )}
              </div>
              <div className="main-formcontainer" style={{ marginTop: "10px" }}>
                <div className="row mb-20">
                  <div className="col-md-6">
                    <label className="font">Voucher Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">{voucherDate}</label>
                  </div>
                  <div className="col-md-6">
                    <label className="font">Voucher Number</label> :
                    &nbsp;&nbsp;
                    <label className="fonttext">{VouchingNumber}</label>
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-6">
                    <label className="font">UTR Date</label> : &nbsp;&nbsp;
                    <label className="fonttext">{UTRDate}</label>
                  </div>
                  <div className="col-md-6">
                    <label className="font">UTR Number</label> : &nbsp;&nbsp;
                    <label className="fonttext">{UTRNumber}</label>
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
                    {attachments.length > 0 && (
                      <ul>
                        {attachments.map((file: any, index: number) => (
                          <li key={index}>
                            <a
                              href={file.ServerRelativeUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                            >
                              {file.Name}
                            </a>
                          </li>
                        ))}
                      </ul>
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
                        {workflowHistory.map((h, index) => (
                          <div key={index} className="history-item">
                            <div>
                              {h.ActionTaken === "Submitted" && "📩 "}
                              {h.ActionTaken === "Approved" && "✅ "}
                              {h.ActionTaken === "Rejected" && "❌ "}
                              {h.ActionTaken === "Send Back" && "↩ "}
                              {h.ActionTaken === "Vouched" && "💰 "}
                              {h.ActionTaken === "Paid" && "💸 "}
                              {h.ActionTaken}
                            </div>

                            <div>
                              <b>{h.CurrentApprover}</b>
                            </div>
                            <div>{h.Comment}</div>
                            <div className="date">
                              {new Date(h.Date).toLocaleString()}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              </div>
              <div className="row my-3">
                <div className="col-md-12">
                  <div className="text-center">
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

export default ViewAdvanceForm;
