# UI5

UI5 related reference  
[Upload data from Excel in CAP Node.js](https://community.sap.com/t5/technology-blogs-by-members/upload-data-from-excel-in-cap-node-js/ba-p/13554121)

A Fiori elements list report is created to display and upload the data into the tables. An upload button is added as an extension to the List Report Application.

The upload entity in this can be re-used for updating data in multiple tables/entities.

The uploaded data is parsed and inserted into the respective entity. We can write validations before upload by internally invoking the custom handlers of the respective entity using `srv.run`.

```js
const query = INSERT.into(entity).entries(data);

let srv = await cds.connect.to('StudentsSrv');

const Result = await srv.run(query);


CAP DB and SRV Artifacts :
entity Students : cuid {
    StudentId: String(6);
    FirstName: String;
    LastName: String;
    DOB: Date;
    Address: String;
}


Service:
service StudentsSrv {
    @cds.persistence.skip
    @odata.singleton
    entity ExcelUpload {
        @Core.MediaType : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        excel : LargeBinary;
    };

    entity Students as projection on db.Students;
}

entity Students as projection on db.Students
}

Service.js: ( Custom Handlers)

srv.on('PUT', "ExcelUpload", async (req, next) => {
    if (req.data.excel) {
        var entity = req.headers.slug;
        const stream = new PassThrough();
        var buffers = [];
        req.data.excel.pipe(stream);
        await new Promise((resolve, reject) => {
            stream.on('data', dataChunk => {
                buffers.push(dataChunk);
            });
            stream.on('end', async () => {
                var buffer = Buffer.concat(buffers);
                var workbook = XLSX.read(buffer, { type: "buffer", cellText: true, cellDates: true, dateNF: 'dd"."mm"."yyyy', cellNF: true, rawNumbers: false });
                let data = []
                const sheets = workbook.SheetNames
                for (let i = 0; i < sheets.length; i++) {
                    const temp = XLSX.utils.sheet_to_json(
                        workbook.Sheets[workbook.SheetNames[i]], { cellText: true, cellDates: true, dateNF: 'dd"."mm"."yyyy', rawNumbers: false })
                    temp.forEach((res, index) => {
                        if (index === 0) return;
                        data.push(JSON.parse(JSON.stringify(res)))
                    })
                }
                if (data) {
                    const responseCall = await CallEntity(entity, data);
                    if (responseCall == -1)
                        reject(req.error(400, JSON.stringify(data)));
                    else {
                        resolve(req.notify({
                            message: 'Upload Successful',
                            status: 200
                        }));
                    }
                }
            });
        });
    } else {
        return next();
    }
});

srv.before('POST', 'Students', async (req) => {
    // Custom validations can be put, if required before upload
});

srv.on('POST', 'Students', async (req) => {
    // Return response to Excel upload entity
});

async function CallEntity(entity, data) {
    if (entity === Students) {
        // If any custom handling is required for a particular entity
    }
    const insertQuery = INSERT.into(entity).entries(data);
    let srv = await cds.connect.to('StudentsSrv');
    const insertResult = await srv.run(insertQuery);
    let query = SELECT.from(entity);
    await srv.run(query);
    return insertResult; // Returns response to Excel upload entity
}


CAP APP Artifacts: Create a list report application using fiori template and extend the application to add custom upload button on the list page.
"controlConfiguration": {
    "@com.sap.vocabularies.UI.v1.LineItem": {
        "actions": {
            "Upload": {
                "id": "UploadButton",
                "text": "Upload",
                "press": "com.test.ExtensionController.Upload",
                "requiresSelection": false
            }
        }
    }
}

ExcelUploadDialog.fragment.xml:
<core:FragmentDefinition xmlns:core="sap.ui.core"
    xmlns:u="sap.ui.unified"
    xmlns="sap.m">
    <Dialog id="_IDDialog"
        title="Excel Upload"
        class="sapUiResponsiveContentPadding"
        beforeOpen=".onBeforeOpen"
        afterClose=".onAfterClose">
        <content>
            <u:FileUploader id="uploader"
                fileType="xlsx"
                multiple="false"
                uploadUrl="StudentsSrv/ExcelUpload/excel"
                fileAllowed=".onFileAllowed"
                fileEmpty=".onFileEmpty"
                uploadComplete=".onUploadComplete"
                typeMissmatch=".onTypeMismatch"
                sendXHR="true"
                useMultipart="false"
                placeholder="Choose a xlsx file..."
                httpRequestMethod="Put" />
        </content>
        <beginButton>
            <Button id="ok"
                text="OK"
                press=".onOk"
                type="Emphasized"
                enabled="false" />
        </beginButton>
        <endButton>
            <Button id="cancel"
                text="Cancel"
                press=".onCancel" />
        </endButton>
    </Dialog>
</core:FragmentDefinition>

ExtensionController:
sap.ui.define(["sap/m/MessageBox", "sap/m/MessageToast", "sap/ui/core/UIComponent"],
    function (MessageBox, MessageToast, UIComponent) {
        "use strict";
        
        function _createUploadController(oExtensionAPI, Entity) {
            var oUploadDialog;

            function setOkButtonEnabled(bOk) {
                oUploadDialog && oUploadDialog.getBeginButton().setEnabled(bOk);
            }

            function setDialogBusy(bBusy) {
                oUploadDialog.setBusy(bBusy)
            }

            function closeDialog() {
                oUploadDialog && oUploadDialog.close()
            }

            function showError(code, target, sMessage) {
                MessageBox.error("Upload failed", {title: "Error"});
            }

            function byId(sId) {
                return sap.ui.core.Fragment.byId("excelUploadDialog", sId);
            }

            return {
                onBeforeOpen: function (oEvent) {
                    oUploadDialog = oEvent.getSource();
                    oExtensionAPI.addDependent(oUploadDialog);
                },

                onAfterClose: function (oEvent) {
                    oExtensionAPI.removeDependent(oUploadDialog);
                    oUploadDialog.destroy();
                    oUploadDialog = undefined;
                },

                onOk: function (oEvent) {
                    setDialogBusy(true);

                    var oFileUploader = byId("uploader");
                    var headPar = new sap.ui.unified.FileUploaderParameter();
                    headPar.setName('slug');
                    headPar.setValue(Entity);
                    oFileUploader.removeHeaderParameter('slug');
                    oFileUploader.addHeaderParameter(headPar);
                    var sUploadUri = oExtensionAPI._controller.extensionAPI._controller._oAppComponent.getManifestObject().resolveUri("./StudentsSrv/ExcelUpload/excel");
                    oFileUploader.setUploadUrl(sUploadUri);
                    oFileUploader
                        .checkFileReadable()
                        .then(function () {
                            oFileUploader.upload();
                        })
                        .catch(function (error) {
                            showError("The file cannot be read.");
                            setDialogBusy(false);
                        })
                },

                onCancel: function (oEvent) {
                    closeDialog();
                },

                onTypeMismatch: function (oEvent) {
                    var sSupportedFileTypes = oEvent
                        .getSource()
                        .getFileType()
                        .map(function (sFileType) {
                            return "*." + sFileType;
                        })
                        .join(", ");

                    showError(
                        "The file type *." +
                        oEvent.getParameter("fileType") +
                        " is not supported. Choose one of the following types: " +
                        sSupportedFileTypes
                    );
                },

                onFileAllowed: function (oEvent) {
                    setOkButtonEnabled(true);
                },

                onFileEmpty: function (oEvent) {
                    setOkButtonEnabled(false);
                },

                onUploadComplete: function (oEvent) {
                    var iStatus = oEvent.getParameter("status");
                    var oFileUploader = oEvent.getSource()

                    oFileUploader.clear();
                    setOkButtonEnabled(false);
                    setDialogBusy(false);

                    if (iStatus >= 400) {
                        var oRawResponse;
                        try {
                            oRawResponse = JSON.parse(oEvent.getParameter("responseRaw"));
                        } catch (e) {
                            oRawResponse = oEvent.getParameter("responseRaw");
                        }
                        if (oRawResponse && oRawResponse.error && oRawResponse.error.message) {
                            showError(oRawResponse.error.code, oRawResponse.error.target, oRawResponse.error.message);
                        }
                    } else {
                        MessageToast.show("File uploaded successfully");
                        oExtensionAPI.refresh();
                        closeDialog();
                    }
                }
            };
        };

        return {
            Upload: function (oBindingContext, aSelectedContexts) {
                this.loadFragment({
                    id: "excelUploadDialog",
                    name: "com.test.ExcelUploadDialog",
                    controller: _createUploadController(this, 'Students')
                }).then(function (oDialog) {
                    oDialog.open();
                });
            }
        };
    });


Summary:

This blog contains code snippets useful for uploading data into tables from excel. We have used a single upload entity which can be re-used for upload of data into various entities. The Upload custom handler parses the uploaded excel using File Uploader and helps to insert that data into respective entity.

References:

https://github.com/SAP-samples/cloud-cap-samples-java/blob/main/app/admin/webapp/extension/Upload.js


    
