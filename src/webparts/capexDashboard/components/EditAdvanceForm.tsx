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
import logo from "../assets/sona-comstarlogo.png";
import "bootstrap/dist/css/bootstrap.min.css";
interface IVendor {
  Id: number;
  VendorCode: string;
  VendorName: string;
}

const EditAdvanceForm = ({ context, formData, onClose }: any) => {
  const sp = spfi().using(SPFx(context));

  // =========================
  // STATES
  // =========================
   const today = new Date();
  const localDate: string = new Date(
  today.getTime() - today.getTimezoneOffset() * 60000
)
  .toISOString()
  .split("T")[0];

  const [previousAdvances, setPreviousAdvances] = useState<any[]>([]);
  const [vendors, setVendors] = useState<IVendor[]>([]);
  const [employee, setEmployee] = useState<any>({});
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [isDraftSaving, setIsDraftSaving] = useState(false);
  const [attachments, setAttachments] = useState<any[]>([]);
  const [selectedFiles, setSelectedFiles] = useState<File[]>([]);
  const [selectedUser, setSelectedUser] = useState<any[]>([]);
  // 🔥 NEW MRN FIELDS
  const [mrnNumber, setMrnNumber] = useState("");
  const [mrnDate, setMrnDate] = useState("");
  const [mrnAmount, setMrnAmount] = useState("");
  const [requestedAmount, setRequestedAmount] = useState("");

  // 🔥 FINAL PAYMENT
  const [finalPayment, setFinalPayment] = useState(false);
  const [installationDetails, setInstallationDetails] = useState("");

  const [selectedVendorId, setSelectedVendorId] = useState<number | null>(null);
  const [selectedVendorName, setSelectedVendorName] = useState("");

  const [poNumber, setPoNumber] = useState("");
  const [poDate, setPoDate] = useState("");
  const [poTerms, setPoTerms] = useState("");
  const [poAmount, setPoAmount] = useState("");

  const [vendorName, setVendorName] = useState("");

  const [approvalMatrix, setApprovalMatrix] = useState<any[]>([]);
  const [workflowHistory, setWorkflowHistory] = useState<any[]>([]);

  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  };

  // =========================
  // LOAD DATA
  // =========================
  const getPreviousAdvances = async (vendorId: number) => {
    try {
      debugger;
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
  const handleNumberChange = (value: string, setter: any) => {
    // Allow only numbers and decimal (max one dot)
    const regex = /^\d*\.?\d*$/;

    if (regex.test(value)) {
      void setter(value);
    }
  };
  const getVendors = async () => {
  try {
    const data = await sp.web.lists
      .getByTitle("VendorMaster")
      .items.select("Id", "VendorCode", "VendorName", "Status")
      .filter("Status eq 'Active'")()
;

    setVendors(data);
  } catch (error) {
    console.error("Vendor fetch error:", error);
  }
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
  const buildApprovalFlow = async () => {
    const flow: any[] = [];

    // 🔹 RM
    if (employee.ReportingManager?.Id) {
      flow.push({
        Id: employee.ReportingManager.Id,
        Name: employee.ReportingManager.Title,
        Role: "RM",
        Level: 1,
        Status: "Pending",
      });
    }

    // 🔹 HOD
    if (employee.HOD?.Id) {
      flow.push({
        Id: employee.HOD.Id,
        Name: employee.HOD.Title,
        Role: "HOD",
        Level: 2,
        Status: "Pending",
      });
    }

    // 🔹 Matrix approvers
    const matrixData = await sp.web.lists
      .getByTitle("CapexPaymentApprovalMatrix")
      .items.select("Role/RoleName,Approver/Id,Approver/Title,Level/Level")
      .expand("Approver,Role,Level")
      .filter("Status eq 'Active'")
      .orderBy("Level", true)();

    const matrixApprovers = matrixData.map((item: any, index: number) => ({
      Id: item.Approver?.Id,
      Name: item.Approver?.Title,
      Role: item.Role?.RoleName,
      Level: flow.length + index + 1,
      Status: "Pending",
    }));

    const finalFlow = [...flow, ...matrixApprovers];

    // 🔥 first approver active
    if (finalFlow.length > 0) {
      finalFlow[0].Status = "In Progress";
    }

    return finalFlow;
  };
  const handledraft = async () => {
    if (isDraftSaving) return;

    setIsDraftSaving(true);

    try {
      debugger;

      const flow = await buildApprovalFlow();
      const currentApprover = flow.length > 0 ? flow[0].Id : null;
      let ensuredUserId: number | null = null;

      // 🔥 preserve history
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        //  Comment: remarks,
        Date: new Date().toISOString(),
      });
      await sp.web.lists
        .getByTitle("CapexPayment")
        .items.getById(formData.ID)
        .update({
          Title: formData.capexId,
          CapexId: formData.capexId,

          // Employee
          EmployeeCode: employee.EmployeeCode,
          EmployeeName: employee.EmployeeName,
          Division: employee.Division,
          Location: employee.Location,
          Email: employee.EmployeeEmail,
          RM: employee.ReportingManager?.Title,
          HOD: employee.HOD?.Title,
          ContactNo: employee.ContactNo,
          EmployeeStatus: employee.EmployeeStatus,

          // Vendor
          VendorCode: selectedVendorId ? selectedVendorId.toString() : "",

          VendorName: selectedVendorName,

          // PO
          PONumber: poNumber,
          PODate: poDate ? new Date(poDate) : null,
          POPaymentTerms: poTerms,
          POAmount: poAmount ? poAmount.toString() : "",

          // 🔥 NEW MRN FIELDS
          MRNNumber: mrnNumber,
          MRNDtae: mrnDate ? new Date(mrnDate) : null,

          MRNAmountwithGST: mrnAmount?.toString(),
          RequestedAmountforPayment: requestedAmount
            ? requestedAmount.toString()
            : "",

          // 🔥 FINAL PAYMENT
          FinalPaymentAgainstPO: finalPayment,

          InstallationDetails: installationDetails,

          // Advance
          // RequestAdvanceAmount: advanceAmount,
          // PaidAmount: paidAmount,

          // PIC

          StatusFlow: "Draft",
          Status: "Draft",
          ApprovalMatrix: JSON.stringify(flow),
          CurrentApproverId: currentApprover,
          WorkflowHistory: JSON.stringify(history),
        });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Draft saved successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=User";
      void handleExit();
      // window.location.reload();
    } catch (error) {
      console.error("ERROR:", error);
      alert("Error while saving ❌");
    } finally {
      setIsDraftSaving(false);
    }
  };

  const handleDeleteExistingFile = async (file: any) => {
    try {
      if (!window.confirm(`Delete ${file.Name}?`)) return;

      await sp.web.getFileByServerRelativePath(file.ServerRelativeUrl).delete();

      // update UI
      setAttachments((prev) =>
        prev.filter((f) => f.ServerRelativeUrl !== file.ServerRelativeUrl),
      );

      alert("File deleted ✅");
    } catch (error) {
      console.error("Delete error:", error);
    }
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

  // =========================
  // UPLOAD FILES
  // =========================

  // =========================
  // VALIDATION
  // =========================

  const handleExit = () => {
    if (onClose) {
      onClose();
    } else {
      window.location.reload();
    }
  };
  const ensureFolder = async (folderPath: string) => {
    try {
      await sp.web.getFolderByServerRelativePath(folderPath)();
    } catch {
      // create folder if not exists
      const parentPath = folderPath.substring(0, folderPath.lastIndexOf("/"));
      const folderName = folderPath.substring(folderPath.lastIndexOf("/") + 1);

      await sp.web
        .getFolderByServerRelativePath(parentPath)
        .folders.addUsingPath(folderName);
    }
  };
  const uploadFiles = async () => {
    try {
      if (!formData?.CapexID || selectedFiles.length === 0) return;

      const safe = formData.CapexID.replace(/\//g, "_");
      const folderPath = `/sites/SonaFinance/CapexPaymentDocs/${safe}`;

      // ✅ Ensure folder exists
      await ensureFolder(folderPath);

      for (const file of selectedFiles) {
        await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, file, { Overwrite: true });
      }

      setSelectedFiles([]);
      await getAttachments(formData.CapexID);
    } catch (error) {
      console.error("Upload error:", error);
      alert("File upload failed ❌");
    }
  };
  const validateForm = () => {
    const errors: string[] = [];

    if (!selectedVendorId) {
      errors.push("Please select Vendor");
    }

    if (!poNumber) {
      errors.push("Please enter PO Number");
    }

    if (!poDate) {
      errors.push("Please select PO Date");
    }
 if (poDate > localDate) {
      errors.push("PO date cannot be a future date");
      // return;
    }
    if (!poTerms) {
      errors.push("Please enter PO Terms");
    }

    if (!poAmount) {
      errors.push("Please enter PO Amount");
    }

    // 🔥 MRN Validation
    if (!mrnNumber) {
      errors.push("Please enter MRN Number");
    }

    if (!mrnDate) {
      errors.push("Please select MRN Date");
    }

    if (!mrnAmount) {
      errors.push("Please enter MRN Amount");
    }

    if (!requestedAmount) {
      errors.push("Please enter Requested Amount");
    }

    // 🔥 Final Payment Validation
    // if (!finalPayment) {
    //   errors.push("Please select Final Payment option");
    // }

    if (finalPayment && !installationDetails) {
      errors.push("Please enter Installation Details");
    }

    if (!attachments || attachments.length === 0) {
      errors.push("Please upload attachment");
    }
    if (!requestedAmount) {
      errors.push("Please enter Requested Amount");
    }

    if (
      requestedAmount &&
      mrnAmount &&
      Number(requestedAmount) > Number(mrnAmount)
    ) {
      errors.push(
        "Requested Amount for Payment should not be greater than MRN Amount (GST)",
      );
    }
    return errors;
  };

  // =========================
  // UPDATE
  // =========================
  const handleSubmit = async () => {
    if (isSubmitting) return;

    setIsSubmitting(true);

    try {
      const errors = validateForm();
      if (errors.length > 0) {
        alert(errors.join("\n"));
        return;
      }

      const ensuredUser = await sp.web.ensureUser(
        selectedUser[0]?.secondaryText,
      );
      // 🔥 preserve matrix (no reset)
      const existingFlow = formData.ApprovalMatrix
        ? JSON.parse(formData.ApprovalMatrix)
        : [];

      // 🔥 preserve history
      const history = formData.WorkFlowHistory
        ? JSON.parse(formData.WorkFlowHistory)
        : [];

      // 🔥 add edit entry
      history.push({
        CurrentApprover: employee.EmployeeName,
        ActionTaken: "Edited",
        // Comment: remarks,
        Date: new Date().toISOString(),
      });
      //const currentApproverId = formData.CurrentApproverId || null;
      await sp.web.lists.getByTitle("CapexPayment").items.add({
        Title: formData.CapexId,
        CapexId: formData.CapexId,

        // Employee
        EmployeeCode: employee.EmployeeCode,
        EmployeeName: employee.EmployeeName,
        Division: employee.Division,
        Location: employee.Location,
        Email: employee.EmployeeEmail,
        RM: employee.ReportingManager?.Title,
        HOD: employee.HOD?.Title,
        ContactNo: employee.ContactNo,
        EmployeeStatus: employee.EmployeeStatus,

        // Vendor
        VendorCode: selectedVendorId ? selectedVendorId.toString() : "",

        VendorName: selectedVendorName,

        // PO
        PONumber: poNumber,
        PODate: poDate ? new Date(poDate) : null,
        POPaymentTerms: poTerms,
        POAmount: poAmount ? poAmount.toString() : "",

        // 🔥 NEW MRN FIELDS
        MRNNumber: mrnNumber,
        MRNDtae: mrnDate ? new Date(mrnDate) : null,

        MRNAmountwithGST: mrnAmount?.toString(),
        RequestedAmountforPayment: requestedAmount
          ? requestedAmount.toString()
          : "",

        // 🔥 FINAL PAYMENT
        FinalPaymentAgainstPO: finalPayment,

        InstallationDetails: installationDetails,

        // Advance
        // RequestAdvanceAmount: advanceAmount,
        // PaidAmount: paidAmount,

        // PIC

        StatusFlow: "Pending for Approver",
        Status: "Pending for Approver",
        ApprovalMatrix: JSON.stringify(existingFlow),
      //  CurrentApproverId: currentApproverId,
        WorkFlowHistory: JSON.stringify(history),
      });

      if (selectedFiles.length > 0) {
        await uploadFiles();
      }

      alert("Updated successfully ✅");
      window.location.href =
        "https://isriglobal.sharepoint.com/sites/SonaFinance/SitePages/CapexPayment.aspx?page=User";
      void handleExit();
    } catch (error: any) {
      console.error("FULL ERROR:", error);
      console.error("SP ERROR:", error?.data?.responseBody);
      alert(error?.data?.responseBody || "Error while saving ❌");
    } finally {
      setIsSubmitting(false);
    }
  };

  // =========================
  // BIND DATA
  // =========================
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
    //setFinalPayment(formData.FinalPaymentAgainstPO ? "Yes" : "No");

    setInstallationDetails(formData.InstallationDetails || "");

    // setApproverRemarks(formData.ApproverRemarks || "");
    // setVoucherDate(formData.VoucherDate?.split("T")[0] || "");
    // setVouchingNumber(formData.VouchingNumber || "");
    // setUTRDate(formData.UTRDate?.split("T")[0] || "");
    // setUTRNumber(formData.UTRNumber || "");

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

  useEffect(() => {
    void getLoggedInUser();
    void getVendors();
    debugger;
    if (selectedVendorId) {
      void getPreviousAdvances(selectedVendorId);
    }
  }, []);

  // =========================
  // UI
  // =========================
  return (
    <div className="MainUplodForm" style={{ margin: "5px 0px" }}>
      <div className="row">
        <div className="col-md-12">
          <div className="Main-Boxpoup">
            {/* 🔹 Header */}
            <div className="bordered">
              <img src={logo} />
              <h1>Edit Advance Payment </h1>
            </div>
            {approvalMatrix.length === 0 ? (
              <p>No approval data</p>
            ) : (
              <div className="displayWF">
                <ul className="approval-flow">
                  {approvalMatrix.map((a, index) => (
                    <li
                      key={index}
                      className={`approval-step ${
                        a.Status === "In Progress"
                          ? "active"
                          : a.Status === "Approved"
                            ? "approved"
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
                <label>Vendor & PO Details</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Vendor Code</label>
                    <span className="required">*</span>
                    <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);

                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");

                        // 🔥 IMPORTANT: Fetch previous advances
                        if (id) {
                          void getPreviousAdvances(id);
                        }
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select>

                    {/* <select
                      value={selectedVendorId || ""}
                      onChange={(e) => {
                        const id = Number(e.target.value);
                        const vendor = vendors.find((v) => v.Id === id);
                        setSelectedVendorId(id);
                        setSelectedVendorName(vendor?.VendorName || "");
                      }}
                      className="formtext-control"
                    >
                      <option value="">Select Vendor</option>
                      {vendors.map((v) => (
                        <option key={v.Id} value={v.Id}>
                          {v.VendorCode}
                        </option>
                      ))}
                    </select> */}
                  </div>
                  <div className="col-md-4">
                    <label className="font">Vendor Name</label>
                    <span className="required">*</span>
                    <input
                      value={selectedVendorName || vendorName}
                      className="form-control readonly"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Number</label>
                    <span className="required">*</span>
                    <input
                      value={poNumber}
                      onChange={(e) => setPoNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>
                </div>
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">PO Date</label>
                    <span className="required">*</span>
                    <input
                      type="date"
                      value={poDate}
                       max={new Date().toISOString().split("T")[0]}
                      onChange={(e) => setPoDate(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Advance Terms</label>
                    <span className="required">*</span>
                    <input
                      value={poTerms}
                      onChange={(e) => setPoTerms(e.target.value)}
                      className="form-control"
                    />
                  </div>
                  <div className="col-md-4">
                    <label className="font">PO Amount (GST)</label>
                    <span className="required">*</span>
                    <input
                      value={poAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setPoAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>MRN Details</label>
              </div>

              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">MRN Number</label>
                    <span className="required">*</span>
                    <input
                      value={mrnNumber}
                      onChange={(e) => setMrnNumber(e.target.value)}
                      className="form-control"
                    />
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Date</label>
                    <span className="required">*</span>
                    <input
                      type="date"
                      value={mrnDate}
                      onChange={(e) => setMrnDate(e.target.value)}
                      className="form-control"
                    />
                  </div>

                  <div className="col-md-4">
                    <label className="font">MRN Amount (GST)</label>
                    <span className="required">*</span>
                    <input
                      value={mrnAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setMrnAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>

                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Requested Amount for Payment</label>
                    <span className="required">*</span>
                    <input
                      value={requestedAmount}
                      onChange={(e) =>
                        handleNumberChange(e.target.value, setRequestedAmount)
                      }
                      className="form-control"
                    />
                  </div>
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Final Payment</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Final Payment Against PO</label>
                    <span className="required">*</span>
                    <select
                      value={finalPayment ? "Yes" : "No"}
                      onChange={(e) =>
                        setFinalPayment(e.target.value === "Yes")
                      }
                      className="formtext-control"
                    >
                      <option value="No">No</option>
                      <option value="Yes">Yes</option>
                    </select>
                  </div>

                  {/* ✅ SHOW ONLY WHEN YES */}
                  {finalPayment && (
                    <div className="col-md-4">
                      <label className="font">Installation Details</label>
                      <span className="required">*</span>
                      <input
                        value={installationDetails}
                        onChange={(e) => setInstallationDetails(e.target.value)}
                        className="form-control"
                      />
                    </div>
                  )}
                </div>
              </div>

              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Previous Advances</label>
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-12">
                    <div style={{ overflowX: "auto" }}>
                      <div className="table-vert-scroll">
                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                          <thead
                            className="text-white"
                            style={{ backgroundColor: "rgb(60, 62, 69)" }}
                          >
                            <tr>
                              <th className="px-4 py-2">PO Number</th>
                              <th className="px-4 py-2">Previous Advance</th>
                              <th className="px-4 py-2">Requested Date</th>
                              <th className="px-4 py-2">Paid Date</th>
                              <th className="px-4 py-2">MRN No</th>
                              <th className="px-4 py-2">Settled Amount</th>
                              <th className="px-4 py-2">Pending Advance</th>
                            </tr>
                          </thead>
                          <tbody>
                            {previousAdvances.length === 0 ? (
                              <tr>
                                <td colSpan={7} style={{ textAlign: "center" }}>
                                 No previous advances available
                                </td>
                              </tr>
                            ) : (
                              previousAdvances.map(
                                (item: any, index: number) => {
                                  const pending = Math.max(
                                    0,
                                    Number(item.RequestAdvanceAmount || 0) -
                                      Number(item.PaidAmount || 0),
                                  );
                                  return (
                                    <tr key={index}>
                                      <td>{item.PONumber}</td>
                                      <td>{item.RequestAdvanceAmount}</td>

                                      <td>
                                        {item.Created
                                          ? new Date(
                                              item.Created,
                                            ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>
                                        {item.VoucherDate
                                          ? new Date(
                                              item.VoucherDate,
                                            ).toLocaleDateString()
                                          : ""}
                                      </td>

                                      <td>{item.VoucherNumber}</td>
                                      <td>{item.PaidAmount}</td>
                                      <td>{pending}</td>
                                    </tr>
                                  );
                                },
                              )
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div className="heading1" style={{ marginTop: "10px" }}>
                <label>Upload Document</label>
                
              </div>
              <div className="main-formcontainer">
                <div className="row mb-20">
                  <div className="col-md-4">
                    <label className="font">Attachments <span className="required">*</span></label>
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

                            {/* 🔥 DELETE FROM SHAREPOINT */}
                            <button
                              type="button"
                              style={{ marginLeft: "10px", color: "red" }}
                              onClick={() => handleDeleteExistingFile(file)}
                            >
                              Delete
                            </button>
                          </li>
                        ))}
                      </ul>
                    )}

                    <input
                      type="file"
                      multiple
                      style={{height :"34px", padding: "3px"}}
                      onChange={(e) => {
                        if (e.target.files) {
                          setSelectedFiles(Array.from(e.target.files));
                        }
                      }}
                    />
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
                              {h.ActionTaken === "Approved" && "✅ "}
                              {h.ActionTaken === "Rejected" && "❌ "}
                              {h.ActionTaken === "Send Back" && "↩ "}
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
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  gap: "10px",
                  margin: "10px",
                }}
              >
                <a
                  onClick={!isSubmitting ? handleSubmit : undefined}
                  className={`submit-btn ${isSubmitting ? "disabled-btn" : ""}`}
                >
                  {isSubmitting ? "Submitting..." : "Submit"}
                </a>

                <a
                  onClick={!isDraftSaving ? handledraft : undefined}
                  className={`Rework-btn ${isDraftSaving ? "disabled-btn" : ""}`}
                >
                  {isDraftSaving ? "Saving..." : "Save as Draft"}
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
  );
};
export default EditAdvanceForm;
