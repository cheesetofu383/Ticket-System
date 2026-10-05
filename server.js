const express = require("express");
const path = require("path");
const XLSX = require("xlsx");
const { google } = require("googleapis");
require("dotenv").config();
const { Resend } = require("resend");
const resend = new Resend(process.env.RESEND_API_KEY);
const app = express();

const PORT = process.env.PORT || 3000;


// ============================================================
// GOOGLE CALENDAR AUTHENTICATION
// ============================================================

const googleKeyFile = path.join(
    __dirname,
    "google-service-account.json"
);

const googleAuth = new google.auth.GoogleAuth({
    keyFile: googleKeyFile,
    scopes: [
        "https://www.googleapis.com/auth/calendar"
    ]
});

const calendar = google.calendar({
    version: "v3",
    auth: googleAuth
});


// ============================================================
// EXCEL DATABASE
// ============================================================

const excelFile = path.join(
    __dirname,
    "database",
    "database.xlsx"
);


// ============================================================
// MIDDLEWARE
// ============================================================

app.use(express.json());

app.use(
    express.static(
        path.join(__dirname, "public")
    )
);
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "staff-login.html"));
});


// ============================================================
// READ EXCEL SHEET
// ============================================================

function getSheetData(sheetName) {

    const workbook = XLSX.readFile(excelFile);

    const worksheet =
        workbook.Sheets[sheetName];

    if (!worksheet) {
        throw new Error(
            `Sheet "${sheetName}" was not found in database.xlsx`
        );
    }

    return XLSX.utils.sheet_to_json(
        worksheet
    );
}


// ============================================================
// GET ALL TICKETS
// ============================================================

app.get("/api/tickets", (req, res) => {

    try {

        const tickets =
            getSheetData("Ticket");

        res.json({
            success: true,
            tickets: tickets
        });

    } catch (error) {

        console.error(
            "Error reading Ticket sheet:"
        );

        console.error(error);

        res.status(500).json({
            success: false,
            message:
                "Could not read the Ticket sheet."
        });

    }

});


// ============================================================
// GET CUSTOMER CREDITS
// ============================================================

app.get(
    "/api/customers/:customerId/credits",
    (req, res) => {

        try {

            const customerId =
                req.params.customerId;

            const customers =
                getSheetData("Customer");

            const customer =
                customers.find(
                    item =>
                        String(
                            item["Customer ID"]
                        ) ===
                        String(customerId)
                );

            if (!customer) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Customer not found."
                });

            }

            res.json({

                success: true,

                customerId:
                    customer["Customer ID"],

                customerName:
                    customer["Name"],

                credits:
                    Number(
                        customer["Credits"] || 0
                    )

            });

        } catch (error) {

            console.error(
                "Error reading customer credits:"
            );

            console.error(error);

            res.status(500).json({
                success: false,
                message:
                    "Could not read customer credits."
            });

        }

    }
);



// CREATE NEW TICKET
app.post("/api/tickets", async (req, res) => {

    try {

        const {
            customerName,
            email,
            issue,
            supportType,
            priority,
            status,
            assignedEngineer,
            appointmentDate,
            appointmentTime,
            appointmentDuration,
            onsiteWorkType,
            appointmentStatus
        } = req.body;


        // ============================================================
        // VALIDATE REQUIRED FIELDS
        // ============================================================

        if (
            !customerName ||
            !email ||
            !issue ||
            !supportType
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Customer Name, Email, Issue and Support Type are required."
            });

        }


        // ============================================================
        // CONVERT ON-SITE DURATION TO HOURS
        // ============================================================

        let finalAppointmentDuration = "";

        if (supportType === "On-site Support") {

            const duration =
                Number(appointmentDuration);

            if (!onsiteWorkType) {

                return res.status(400).json({
                    success: false,
                    message:
                        "On-site Work Type is required."
                });

            }

            if (!duration || duration <= 0) {

                return res.status(400).json({
                    success: false,
                    message:
                        "A valid appointment duration is required."
                });

            }


            // --------------------------------------------------------
            // AD HOC
            // Duration is already in hours
            // Minimum 2 hours
            // --------------------------------------------------------

            if (onsiteWorkType === "Ad Hoc") {

                if (duration < 2) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Ad Hoc work requires a minimum of 2 hours."
                    });

                }

                finalAppointmentDuration = duration;

            }


            // --------------------------------------------------------
            // MAINTENANCE
            // Duration is already in hours
            // Minimum 1 hour
            // --------------------------------------------------------

            else if (onsiteWorkType === "Maintenance") {

                if (duration < 1) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Maintenance work requires a minimum of 1 hour."
                    });

                }

                finalAppointmentDuration = duration;

            }


            // --------------------------------------------------------
            // PROJECT
            // Duration is entered in DAYS
            // 1 day = 8 working hours
            // --------------------------------------------------------

            else if (onsiteWorkType === "Project") {

                if (
                    duration < 1 ||
                    !Number.isInteger(duration)
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Project duration must be at least 1 whole day."
                    });

                }

                finalAppointmentDuration =
                    duration * 8;

            }


            else {

                return res.status(400).json({
                    success: false,
                    message:
                        "Invalid On-site Work Type."
                });

            }

        }


        // ============================================================
        // READ WORKBOOK
        // ============================================================

        const workbook =
            XLSX.readFile(excelFile);


        // ============================================================
        // GET TICKET SHEET
        // ============================================================

        const worksheet =
            workbook.Sheets["Ticket"];

        if (!worksheet) {

            throw new Error(
                'Sheet "Ticket" was not found.'
            );

        }


        // ============================================================
        // GET CUSTOMER SHEET
        // ============================================================

        const customerSheet =
            workbook.Sheets["Customer"];

        if (!customerSheet) {

            return res.status(500).json({
                success: false,
                message:
                    "Customer sheet not found."
            });

        }


        // ============================================================
        // READ CUSTOMERS
        // ============================================================

        const customers =
            XLSX.utils.sheet_to_json(
                customerSheet
            );


        // ============================================================
        // NORMALIZE CUSTOMER INFORMATION
        // ============================================================

        const enteredName =
            String(customerName)
                .trim()
                .toLowerCase();

        const enteredEmail =
            String(email)
                .trim()
                .toLowerCase();


        // ============================================================
        // FIND CUSTOMER BY NAME + EMAIL
        // ============================================================

        const customer =
            customers.find(
                customer => {

                    const storedName =
                        String(
                            customer["Name"] || ""
                        )
                            .trim()
                            .toLowerCase();

                    const storedEmail =
                        String(
                            customer["Email"] || ""
                        )
                            .trim()
                            .toLowerCase();

                    return (
                        storedName === enteredName &&
                        storedEmail === enteredEmail
                    );

                }
            );


        // ============================================================
        // CUSTOMER DOES NOT EXIST
        // ============================================================

        if (!customer) {

            return res.status(404).json({
                success: false,
                message:
                    "Customer name and email do not match an existing customer."
            });

        }


        // ============================================================
        // GET VERIFIED CUSTOMER INFORMATION
        // ============================================================

        const realCustomerId =
            customer["Customer ID"];

        const realCustomerName =
            customer["Name"];

        const realEmail =
            customer["Email"];

        const credits =
            Number(
                customer["Credits"] || 0
            );


        // ============================================================
        // MAKE SURE CUSTOMER RECORD IS COMPLETE
        // ============================================================

        if (
            !realCustomerId ||
            !realCustomerName ||
            !realEmail
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Customer record is missing Customer ID, Name or Email."
            });

        }


        // ============================================================
        // READ EXISTING TICKETS
        // ============================================================

        const tickets =
            XLSX.utils.sheet_to_json(
                worksheet
            );


        // ============================================================
        // CREATE TICKET ID
        // ============================================================

        const ticketId =
            `T-${Date.now()}`;


        // ============================================================
        // CREATE NEW TICKET
        // ============================================================

        const newTicket = {

            "Ticket ID":
                ticketId,

            "Customer ID":
                realCustomerId,

            "Customer Name":
                realCustomerName,

            "Email":
                realEmail,

            "Issue":
                issue,

            "Support Type":
                supportType,

            "Priority":
                priority || "Medium",

            "Status":
                status || "Open",

            "Assigned Engineer":
                assignedEngineer || "",

            "Appointment Date":
                appointmentDate || "",

            "Appointment Time":
                appointmentTime || "",

            "On-site Work Type":
                supportType === "On-site Support"
                    ? onsiteWorkType || ""
                    : "",

            "Appointment Duration":
                finalAppointmentDuration,

            "Appointment Status":
                appointmentStatus || "Pending",

            "Created Date":
                new Date().toISOString(),

            "Service Result":
                "",

            "Calendar Event ID":
                "",

            "Credits Charged":
                0,

            "Employee Sign In":
                "",

            "Employee Sign Out":
                "",

            "Work Duration":
                ""

        };


        // ============================================================
        // ADD TICKET
        // ============================================================

        tickets.push(
            newTicket
        );


        // ============================================================
        // SAVE TICKET SHEET
        // ============================================================

        const newWorksheet =
            XLSX.utils.json_to_sheet(
                tickets
            );

        workbook.Sheets["Ticket"] =
            newWorksheet;


        XLSX.writeFile(
            workbook,
            excelFile
        );


        // ============================================================
        // SUCCESS RESPONSE
        // ============================================================

        res.status(201).json({

            success: true,

            message:
                "Ticket created successfully.",

            ticketId:
                ticketId,

            customer: {

                customerId:
                    realCustomerId,

                name:
                    realCustomerName,

                email:
                    realEmail,

                credits:
                    credits

            }

        });

    } catch (error) {

        console.error(
            "Error creating ticket:"
        );

        console.error(error);

        res.status(500).json({

            success: false,

            message:
                "Could not create ticket.",

            error:
                error.message

        });

    }

});
async function deleteCalendarEvent(calendarId, eventId) {
    if (!calendarId || !eventId) {
        return;
    }

    try {
        await calendar.events.delete({
            calendarId: calendarId,
            eventId: eventId
        });

        console.log(`Deleted old calendar event ${eventId}`);
    } catch (error) {
        // Event already doesn't exist — that's okay
        if (
            error.code === 404 ||
            error.response?.status === 404
        ) {
            console.log("Old calendar event was already deleted.");
            return;
        }

        throw error;
    }
}

// ============================================================
// CHECK GOOGLE CALENDAR AVAILABILITY
// ============================================================

async function checkCalendarAvailability(
    calendarId,
    appointmentDate,
    appointmentTime,
    appointmentDuration = 1,
    existingEventId = ""
) {

    const startDateTime =
        `${appointmentDate}T${appointmentTime}:00+08:00`;

    const startTime =
        new Date(startDateTime);

    const duration =
        Number(appointmentDuration);

    const endTime =
        new Date(
            startTime.getTime() +
            duration * 60 * 60 * 1000
        );


    const response =
        await calendar.events.list({

            calendarId:
                calendarId,

            timeMin:
                startTime.toISOString(),

            timeMax:
                endTime.toISOString(),

            singleEvents:
                true,

            orderBy:
                "startTime"

        });


    const events =
        response.data.items || [];


    const conflictingEvents =
        events.filter(event => {

            // Ignore this ticket's own event
            if (
                existingEventId &&
                String(event.id) ===
                String(existingEventId)
            ) {

                return false;

            }


            // Ignore cancelled events
            if (
                event.status ===
                "cancelled"
            ) {

                return false;

            }


            return true;

        });


    return {

        available:
            conflictingEvents.length === 0,

        busyPeriods:
            conflictingEvents.map(
                event => ({

                    start:
                        event.start?.dateTime ||
                        event.start?.date ||
                        null,

                    end:
                        event.end?.dateTime ||
                        event.end?.date ||
                        null,

                    eventId:
                        event.id,

                    summary:
                        event.summary ||
                        "Busy"

                })
            )

    };

}

// EMPLOYEE SIGN IN

app.post(
    "/api/tickets/:ticketId/sign-in",
    (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;

            const workbook =
                XLSX.readFile(excelFile);

            const worksheet =
                workbook.Sheets["Ticket"];

            if (!worksheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Ticket sheet not found."
                });

            }

            const tickets =
                XLSX.utils.sheet_to_json(
                    worksheet
                );

            const ticketIndex =
                tickets.findIndex(
                    ticket =>
                        String(
                            ticket["Ticket ID"]
                        ) ===
                        String(ticketId)
                );

            if (ticketIndex === -1) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }

            const ticket =
                tickets[ticketIndex];


            // ============================================================
            // ONLY REMOTE / ON-SITE SUPPORT
            // ============================================================

            const supportType =
                String(
                    ticket["Support Type"] || ""
                )
                    .trim()
                    .toLowerCase();

            const isWorkSession =
                supportType.includes("remote") ||
                supportType.includes("on-site") ||
                supportType.includes("on site");

            if (!isWorkSession) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Employee work sessions are only available for Remote Support and On-site Support."
                });

            }


            // ============================================================
            // PREVENT DOUBLE SIGN-IN
            // ============================================================

            if (ticket["Employee Sign In"]) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Employee has already signed in."
                });

            }


            // ============================================================
            // SIGN IN TIME
            // ============================================================

            const signInTime =
                new Date().toISOString();

            ticket["Employee Sign In"] =
                signInTime;

            ticket["Employee Sign Out"] =
                "";

            ticket["Work Duration"] =
                "";


            tickets[ticketIndex] =
                ticket;


            workbook.Sheets["Ticket"] =
                XLSX.utils.json_to_sheet(
                    tickets
                );

            XLSX.writeFile(
                workbook,
                excelFile
            );


            res.json({

                success: true,

                message:
                    "Employee signed in successfully.",

                signInTime:
                    signInTime

            });

        } catch (error) {

            console.error(
                "Error signing in:"
            );

            console.error(error);

            res.status(500).json({

                success: false,

                message:
                    "Could not sign in."

            });

        }

    }
);

// EMPLOYEE SIGN OUT

app.post(
    "/api/tickets/:ticketId/sign-out",
    (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;

            const workbook =
                XLSX.readFile(excelFile);

            const worksheet =
                workbook.Sheets["Ticket"];

            if (!worksheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Ticket sheet not found."
                });

            }

            const tickets =
                XLSX.utils.sheet_to_json(
                    worksheet
                );

            const ticketIndex =
                tickets.findIndex(
                    ticket =>
                        String(
                            ticket["Ticket ID"]
                        ) ===
                        String(ticketId)
                );

            if (ticketIndex === -1) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }

            const ticket =
                tickets[ticketIndex];


            // ============================================================
            // CHECK SIGN IN
            // ============================================================

            if (!ticket["Employee Sign In"]) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Employee must sign in before signing out."
                });

            }


            // ============================================================
            // PREVENT DOUBLE SIGN-OUT
            // ============================================================

            if (ticket["Employee Sign Out"]) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Employee has already signed out."
                });

            }


            // ============================================================
            // CALCULATE WORK DURATION
            // ============================================================

            const signInTime =
                new Date(
                    ticket["Employee Sign In"]
                );

            const signOutTime =
                new Date();


            if (
                isNaN(
                    signInTime.getTime()
                )
            ) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Invalid employee sign-in time."
                });

            }


            const durationMilliseconds =
                signOutTime.getTime() -
                signInTime.getTime();


            const durationHours =
                durationMilliseconds /
                (1000 * 60 * 60);


            const workDuration =
                Math.round(
                    durationHours * 100
                ) / 100;


            // ============================================================
            // SAVE
            // ============================================================

            ticket["Employee Sign Out"] =
                signOutTime.toISOString();

            ticket["Work Duration"] =
                workDuration;


            tickets[ticketIndex] =
                ticket;


            workbook.Sheets["Ticket"] =
                XLSX.utils.json_to_sheet(
                    tickets
                );

            XLSX.writeFile(
                workbook,
                excelFile
            );


            res.json({

                success: true,

                message:
                    "Employee signed out successfully.",

                signOutTime:
                    signOutTime.toISOString(),

                workDuration:
                    workDuration

            });

        } catch (error) {

            console.error(
                "Error signing out:"
            );

            console.error(error);

            res.status(500).json({

                success: false,

                message:
                    "Could not sign out."

            });

        }

    }
);
// UPDATE TICKET
app.put(
    "/api/tickets/:ticketId",
    async (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;


            const {
    assignedEngineer,
    status,
    appointmentDuration,
    onsiteWorkType,
    appointmentStatus,
    serviceResult
} = req.body;


            // ============================================================
            // READ EXCEL
            // ============================================================

            const workbook =
                XLSX.readFile(excelFile);

            const worksheet =
                workbook.Sheets["Ticket"];

            if (!worksheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Ticket sheet not found."
                });

            }


            const tickets =
                XLSX.utils.sheet_to_json(
                    worksheet
                );


            // ============================================================
            // FIND TICKET
            // ============================================================

            const ticketIndex =
                tickets.findIndex(
                    ticket =>
                        String(
                            ticket["Ticket ID"]
                        ) ===
                        String(ticketId)
                );


            if (ticketIndex === -1) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }


            const ticket =
                tickets[ticketIndex];

             const effectiveOnsiteWorkType = onsiteWorkType || ticket["On-site Work Type"] || ""; 
            // ============================================================
            // GET EXISTING APPOINTMENT INFORMATION
            // ============================================================

            const appointmentDate =
                ticket["Appointment Date"] || "";

            const appointmentTime =
                ticket["Appointment Time"] || "";


            // ============================================================
            // GET OLD APPOINTMENT STATUS
            // ============================================================

            const oldAppointmentStatus =
                String(
                    ticket["Appointment Status"] ||
                    "Pending"
                ).trim();


            const newAppointmentStatus =
                String(
                    appointmentStatus ||
                    "Pending"
                ).trim();


            console.log(
                "========== APPOINTMENT STATUS =========="
            );

            console.log(
                "Ticket ID:",
                ticketId
            );

            console.log(
                "Old Appointment Status:",
                oldAppointmentStatus
            );

            console.log(
                "New Appointment Status:",
                newAppointmentStatus
            );

            console.log(
                "========================================"
            );


            // ============================================================
            // CREDIT SYSTEM
            // ============================================================

            const supportType =
                String(
                    ticket["Support Type"] || ""
                )
                    .trim()
                    .toLowerCase();


            const oldCreditsCharged =
                Number(
                    ticket["Credits Charged"] || 0
                );

const enteredDuration =
    Number(
        appointmentDuration || 0
    );
let newDuration =
    enteredDuration;


            let newCreditsCharged =
                oldCreditsCharged;


            const isCreditSupport =
                supportType.includes("remote") ||
                supportType.includes("on-site") ||
                supportType.includes("on site");


            // ============================================================
            // REMOTE / ON-SITE SUPPORT
            // ============================================================

            if (isCreditSupport) {

               if (
    !Number.isFinite(
        enteredDuration
    ) ||
    enteredDuration <= 0
) {

    return res.status(400).json({
        success: false,
        message:
            "Appointment Duration must be greater than 0."
    });

}


// ============================================================
// ON-SITE WORK TYPE VALIDATION
// ============================================================

if (
    supportType.includes("on-site") ||
    supportType.includes("on site")
) {

    if (!onsiteWorkType) {

        return res.status(400).json({
            success: false,
            message:
                "On-site Work Type is required."
        });

    }


    // --------------------------------------------------------
    // AD HOC
    // Minimum 2 hours, 0.5 increments
    // --------------------------------------------------------

    if (
        onsiteWorkType === "Ad Hoc"
    ) {

        if (
            newDuration < 2 ||
            newDuration % 0.5 !== 0
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Ad Hoc work requires a minimum of 2 hours and uses 0.5 hour increments."
            });

        }

    }


    // --------------------------------------------------------
    // MAINTENANCE
    // Minimum 1 hour, 0.5 increments
    // --------------------------------------------------------

    else if (
        onsiteWorkType === "Maintenance"
    ) {

        if (
            newDuration < 1 ||
            newDuration % 0.5 !== 0
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Maintenance work requires a minimum of 1 hour and uses 0.5 hour increments."
            });

        }

    }

    else if (
    onsiteWorkType === "Project"
) {

    // Project duration has already been
    // converted from days to hours
    // by the frontend.

    if (
        !Number.isFinite(newDuration) ||
        newDuration <= 0 ||
        newDuration % 8 !== 0
    ) {

        return res.status(400).json({
            success: false,
            message:
                "Project duration must represent a whole number of 8-hour days."
        });

    }

}


    else {

        return res.status(400).json({
            success: false,
            message:
                "Invalid On-site Work Type."
        });

    }

}

                // ========================================================
                // CURRENT TICKET CREDIT CHARGE
                // ========================================================

                newCreditsCharged =
                    newDuration;


                // ========================================================
                // DIFFERENCE FROM PREVIOUS CHARGE
                // ========================================================

                const creditDifference =
                    newCreditsCharged -
                    oldCreditsCharged;


                console.log(
                    "========== CREDIT DEBUG =========="
                );

                console.log(
                    "Ticket ID:",
                    ticketId
                );

                console.log(
                    "Customer ID:",
                    ticket["Customer ID"]
                );

                console.log(
                    "Support Type:",
                    ticket["Support Type"]
                );

                console.log(
                    "Appointment Duration:",
                    newDuration
                );

                console.log(
                    "Old Credits Charged:",
                    oldCreditsCharged
                );

                console.log(
                    "New Credits Charged:",
                    newCreditsCharged
                );

                console.log(
                    "Credit Difference:",
                    creditDifference
                );

                console.log(
                    "=================================="
                );


                // ========================================================
                // ONLY CHANGE CUSTOMER CREDITS IF NEEDED
                // ========================================================

                if (
                    creditDifference !== 0
                ) {

                    const customerSheet =
                        workbook.Sheets["Customer"];


                    if (!customerSheet) {

                        return res.status(500).json({
                            success: false,
                            message:
                                "Customer sheet not found."
                        });

                    }


                    const customers =
                        XLSX.utils.sheet_to_json(
                            customerSheet
                        );


                    // ====================================================
                    // FIND CUSTOMER
                    // ====================================================

                    const customerIndex =
                        customers.findIndex(
                            customer =>
                                String(
                                    customer[
                                        "Customer ID"
                                    ]
                                ) ===
                                String(
                                    ticket[
                                        "Customer ID"
                                    ]
                                )
                        );


                    if (
                        customerIndex === -1
                    ) {

                        return res.status(404).json({
                            success: false,
                            message:
                                "Customer for this ticket could not be found."
                        });

                    }


                    const customer =
                        customers[
                            customerIndex
                        ];


                    const currentCredits =
                        Number(
                            customer[
                                "Credits"
                            ] || 0
                        );


                    // ====================================================
                    // CHECK AVAILABLE CREDITS
                    // ====================================================

                    if (
                        creditDifference > 0 &&
                        currentCredits <
                        creditDifference
                    ) {

                        return res.status(400).json({

                            success: false,

                            message:
                                `Customer does not have enough credits. Required: ${creditDifference}, Available: ${currentCredits}.`,

                            requiredCredits:
                                creditDifference,

                            availableCredits:
                                currentCredits

                        });

                    }


                    // ====================================================
                    // DEDUCT OR REFUND
                    // ====================================================

                    customer["Credits"] =
                        currentCredits -
                        creditDifference;


                    console.log(
                        "Customer Credits BEFORE:",
                        currentCredits
                    );

                    console.log(
                        "Customer Credits AFTER:",
                        customer["Credits"]
                    );

                    console.log(
                        "=================================="
                    );


                    customers[
                        customerIndex
                    ] =
                        customer;


                    workbook.Sheets["Customer"] =
                        XLSX.utils.json_to_sheet(
                            customers
                        );

                }

            } else {

                // ========================================================
                // PHONE / EMAIL SUPPORT
                // ========================================================

                newCreditsCharged =
                    0;

            }


            // ============================================================
            // UPDATE TICKET INFORMATION
            // ============================================================

            ticket["Assigned Engineer"] =
                assignedEngineer || "";


            ticket["Status"] =
                status || "Open";


            ticket["Appointment Duration"] = newDuration;

            ticket["On-site Work Type"] =
    effectiveOnsiteWorkType;

            ticket["Appointment Status"] =
                newAppointmentStatus;


            ticket["Service Result"] =
                serviceResult || "";


            ticket["Credits Charged"] =
                newCreditsCharged;


            // ============================================================
            // GOOGLE CALENDAR
            // ============================================================

            let calendarEventId =
                ticket["Calendar Event ID"] ||
                "";


            // ============================================================
            // IMPORTANT APPOINTMENT STATUS RULE
            // ============================================================
            //
            // PENDING
            // Appointment exists but is NOT confirmed.
            //
            // CONFIRMED
            // Appointment should exist on Google Calendar.
            //
            // RESCHEDULE REQUIRED
            // Do NOT change the current appointment yet.
            //
            // ============================================================


            if (
                newAppointmentStatus ===
                "Reschedule Required"
            ) {

                console.log(
                    "Appointment marked as Reschedule Required."
                );

                console.log(
                    "Existing date/time will remain unchanged."
                );

                // --------------------------------------------------------
                // DO NOT CREATE OR UPDATE GOOGLE CALENDAR HERE.
                // The existing appointment information stays in Excel.
                // Actual rescheduling will be handled separately.
                // --------------------------------------------------------

            }


            // ============================================================
            // CONFIRMED APPOINTMENT
            // ============================================================

            else if (
                newAppointmentStatus ===
                "Confirmed" &&
                assignedEngineer &&
                appointmentDate &&
                appointmentTime &&
                appointmentDuration
            ) {

                // ========================================================
                // READ STAFF SHEET
                // ========================================================

                const staffSheet =
                    workbook.Sheets["Staff"];


                if (!staffSheet) {

                    return res.status(500).json({
                        success: false,
                        message:
                            "Staff sheet not found."
                    });

                }


                const staff =
                    XLSX.utils.sheet_to_json(
                        staffSheet
                    );


                // ========================================================
                // FIND ASSIGNED ENGINEER
                // ========================================================

                const engineer =
                    staff.find(
                        person =>
                            String(
                                person["Name"] ||
                                ""
                            )
                                .trim()
                                .toLowerCase() ===
                            String(
                                assignedEngineer
                            )
                                .trim()
                                .toLowerCase()
                    );


                if (!engineer) {

                    return res.status(404).json({
                        success: false,
                        message:
                            "Assigned engineer could not be found."
                    });

                }


                // ========================================================
                // GET ENGINEER CALENDAR ID
                // ========================================================

                const calendarId =
                    String(
                        engineer["Calendar ID"] ||
                        ""
                    ).trim();


                if (!calendarId) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Assigned engineer does not have a Calendar ID."
                    });

                }


                // ========================================================
                // VALIDATE DURATION
                // ========================================================

                const duration =
                    Number(
                        appointmentDuration
                    );


                if (
                    !Number.isFinite(
                        duration
                    ) ||
                    duration <= 0 ||
                    duration % 0.5 !== 0
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Appointment Duration must be greater than 0 and use 0.5 hour increments."
                    });

                }

                // CHECK ENGINEER CALENDAR

                const availability =
                    await checkCalendarAvailability(
                        calendarId,
                        appointmentDate,
                        appointmentTime,
                        duration,
                        calendarEventId
                    );


                if (
                    !availability.available
                ) {

                    return res.status(409).json({

                        success: false,

                        message:
                            "The engineer is not available at the selected appointment time.",

                        engineer:
                            engineer["Name"],

                        calendarId:
                            calendarId,

                        appointmentDate:
                            appointmentDate,

                        appointmentTime:
                            appointmentTime,

                        appointmentDuration:
                            duration,

                        busyPeriods:
                            availability.busyPeriods

                    });

                }


                // ========================================================
                // UPDATE EXISTING GOOGLE CALENDAR EVENT
                // ========================================================

                if (calendarEventId) {

                    console.log(
                        "Updating existing Google Calendar event..."
                    );

                    console.log(
                        "Calendar ID:",
                        calendarId
                    );

                    console.log(
                        "Event ID:",
                        calendarEventId
                    );

                    console.log(
                        "Appointment Date:",
                        appointmentDate
                    );

                    console.log(
                        "Appointment Time:",
                        appointmentTime
                    );

                    console.log(
                        "New Duration:",
                        duration
                    );


                    await updateCalendarEventDuration(
                        calendarId,
                        calendarEventId,
                        appointmentDate,
                        appointmentTime,
                        duration
                    );


                    console.log(
                        "Google Calendar event updated successfully."
                    );

                }

                // ========================================================
                // CREATE NEW GOOGLE CALENDAR EVENT
                // ========================================================

                else {

                    console.log(
                        "No Calendar Event ID found."
                    );

                    console.log(
                        "Creating new Google Calendar event..."
                    );


                    const googleEvent =
                        await createCalendarEvent(
                            calendarId,
                            ticketId,
                            ticket["Customer Name"],
                            ticket["Issue"],
                            engineer["Name"],
                            appointmentDate,
                            appointmentTime,
                            duration
                        );


                    calendarEventId =
                        googleEvent.id;


                    ticket[
                        "Calendar Event ID"
                    ] =
                        calendarEventId;


                    console.log(
                        "New Calendar Event ID:",
                        calendarEventId
                    );

                }

            }


            // ============================================================
            // SAVE TICKET
            // ============================================================

            tickets[ticketIndex] =
                ticket;


            const newWorksheet =
                XLSX.utils.json_to_sheet(
                    tickets
                );


            workbook.Sheets["Ticket"] =
                newWorksheet;


            XLSX.writeFile(
                workbook,
                excelFile
            );


            // ============================================================
            // RESPONSE
            // ============================================================

            res.json({

                success: true,

                message:
                    "Ticket updated successfully.",

                ticketId:
                    ticketId,

                appointmentStatus:
                    newAppointmentStatus,

                creditsCharged:
                    newCreditsCharged,

                appointmentDuration:
                    appointmentDuration,

                calendarEventId:
                    calendarEventId || ""

            });


        } catch (error) {

            console.error(
                "Error updating ticket:"
            );

            console.error(error);


            res.status(500).json({

                success: false,

                message:
                    "Could not update ticket.",

                error:
                    error.response?.data ||
                    error.message

            });

        }

    }
);


// DELETE TICKET
app.delete(
    "/api/tickets/:ticketId",
    async (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;


            // ============================================================
            // READ EXCEL
            // ============================================================

            const workbook =
                XLSX.readFile(excelFile);

            const worksheet =
                workbook.Sheets["Ticket"];


            if (!worksheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Ticket sheet not found."
                });

            }


            const tickets =
                XLSX.utils.sheet_to_json(
                    worksheet
                );


            // ============================================================
            // FIND TICKET
            // ============================================================

            const ticketIndex =
                tickets.findIndex(
                    ticket =>
                        String(
                            ticket["Ticket ID"]
                        ) ===
                        String(ticketId)
                );


            if (ticketIndex === -1) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }


            const ticket =
                tickets[ticketIndex];

            // ============================================================
// REFUND CUSTOMER CREDITS
// ============================================================

const creditsCharged =
    Number(
        ticket["Credits Charged"] || 0
    );


if (
    creditsCharged > 0
) {

    const customerSheet =
        workbook.Sheets["Customer"];


    if (!customerSheet) {

        return res.status(500).json({
            success: false,
            message:
                "Customer sheet not found."
        });

    }


    const customers =
        XLSX.utils.sheet_to_json(
            customerSheet
        );


    const customerIndex =
        customers.findIndex(
            customer =>
                String(
                    customer["Customer ID"]
                ).trim() ===
                String(
                    ticket["Customer ID"]
                ).trim()
        );


    if (
        customerIndex === -1
    ) {

        return res.status(404).json({
            success: false,
            message:
                "Customer for this ticket could not be found."
        });

    }


    const customer =
        customers[customerIndex];


    const currentCredits =
        Number(
            customer["Credits"] || 0
        );


    customer["Credits"] =
        currentCredits +
        creditsCharged;


    console.log(
        "========== CREDIT REFUND =========="
    );

    console.log(
        "Ticket ID:",
        ticketId
    );

    console.log(
        "Customer ID:",
        ticket["Customer ID"]
    );

    console.log(
        "Credits Charged:",
        creditsCharged
    );

    console.log(
        "Credits BEFORE:",
        currentCredits
    );

    console.log(
        "Credits AFTER:",
        customer["Credits"]
    );

    console.log(
        "===================================="
    );


    customers[customerIndex] =
        customer;


    workbook.Sheets["Customer"] =
        XLSX.utils.json_to_sheet(
            customers
        );

}
            // ============================================================
            // GET CALENDAR EVENT ID
            // ============================================================

            const calendarEventId =
                String(
                    ticket["Calendar Event ID"] || ""
                ).trim();


            // ============================================================
            // DELETE GOOGLE CALENDAR EVENT
            // ============================================================

            if (calendarEventId) {

                console.log(
                    "========== DELETE CALENDAR EVENT =========="
                );

                console.log(
                    "Ticket ID:",
                    ticketId
                );

                console.log(
                    "Calendar Event ID:",
                    calendarEventId
                );


                // --------------------------------------------------------
                // Find the assigned engineer
                // --------------------------------------------------------

                const staffSheet =
                    workbook.Sheets["Staff"];


                if (!staffSheet) {

                    return res.status(500).json({
                        success: false,
                        message:
                            "Staff sheet not found."
                    });

                }


                const staff =
                    XLSX.utils.sheet_to_json(
                        staffSheet
                    );


                const assignedEngineer =
                    String(
                        ticket["Assigned Engineer"] || ""
                    ).trim();


                const engineer =
                    staff.find(
                        person =>
                            String(
                                person["Name"] || ""
                            )
                                .trim()
                                .toLowerCase() ===
                            assignedEngineer.toLowerCase()
                    );


                if (!engineer) {

                    return res.status(404).json({
                        success: false,
                        message:
                            "Assigned engineer could not be found."
                    });

                }


                // --------------------------------------------------------
                // Get engineer Calendar ID
                // --------------------------------------------------------

                const calendarId =
                    String(
                        engineer["Calendar ID"] || ""
                    ).trim();


                if (!calendarId) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Assigned engineer does not have a Calendar ID."
                    });

                }


                // --------------------------------------------------------
                // Delete Google Calendar event
                // --------------------------------------------------------

                await calendar.events.delete({

                    calendarId:
                        calendarId,

                    eventId:
                        calendarEventId

                });


                console.log(
                    "Google Calendar event deleted successfully."
                );

                console.log(
                    "==========================================="
                );

            }


            // ============================================================
            // REMOVE TICKET FROM EXCEL
            // ============================================================

            tickets.splice(
                ticketIndex,
                1
            );


            // ============================================================
            // SAVE EXCEL
            // ============================================================

            const newWorksheet =
                XLSX.utils.json_to_sheet(
                    tickets
                );


            workbook.Sheets["Ticket"] =
                newWorksheet;


            XLSX.writeFile(
                workbook,
                excelFile
            );


            // ============================================================
            // SUCCESS
            // ============================================================

            res.json({

                success: true,

                message:
                    "Ticket and calendar appointment deleted successfully.",

                ticketId:
                    ticketId

            });


        } catch (error) {

            console.error(
                "Error deleting ticket:"
            );

            console.error(error);


            res.status(500).json({

                success: false,

                message:
                    "Could not delete ticket.",

                error:
                    error.response?.data ||
                    error.message

            });

        }

    }
);

// ============================================================
// GET ONE TICKET
// ============================================================

app.get(
    "/api/tickets/:ticketId",
    (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;


            const tickets =
                getSheetData("Ticket");


            const ticket =
                tickets.find(
                    item =>
                        String(
                            item["Ticket ID"]
                        ) ===
                        String(ticketId)
                );


            if (!ticket) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }


            res.json({

                success: true,

                ticket:
                    ticket

            });


        } catch (error) {

            console.error(
                "Error reading ticket:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    "Could not read the ticket."

            });

        }

    }
);


// ============================================================
// GET ALL STAFF
// ============================================================

app.get(
    "/api/staff",
    (req, res) => {

        try {

            const staff =
                getSheetData("Staff");


            res.json({

                success: true,

                staff:
                    staff

            });


        } catch (error) {

            console.error(
                "Error reading Staff sheet:"
            );

            console.error(error);


            res.status(500).json({

                success: false,

                message:
                    "Could not read the Staff sheet."

            });

        }

    }
);


// ============================================================
// TEST GOOGLE CALENDAR ACCESS DYNAMICALLY
// ============================================================

app.get(
    "/api/test-calendar",
    async (req, res) => {

        try {

            const engineerName =
                req.query.engineerName;


            if (!engineerName) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Engineer name is required."
                });

            }


            // ========================================================
            // READ STAFF
            // ========================================================

            const workbook =
                XLSX.readFile(excelFile);


            const staffSheet =
                workbook.Sheets["Staff"];


            if (!staffSheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Staff sheet not found."
                });

            }


            const staff =
                XLSX.utils.sheet_to_json(
                    staffSheet
                );


            // ========================================================
            // FIND ENGINEER
            // ========================================================

            const engineer =
                staff.find(
                    person =>
                        String(
                            person["Name"] ||
                            ""
                        )
                            .trim()
                            .toLowerCase() ===
                        String(
                            engineerName
                        )
                            .trim()
                            .toLowerCase()
                );


            if (!engineer) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Engineer could not be found."
                });

            }


            // ========================================================
            // GET CALENDAR ID
            // ========================================================

            const calendarId =
                String(
                    engineer["Calendar ID"] ||
                    ""
                ).trim();


            if (!calendarId) {

                return res.status(400).json({
                    success: false,
                    message:
                        "This engineer does not have a Calendar ID."
                });

            }


            // ========================================================
            // TEST GOOGLE CALENDAR
            // ========================================================

            const response =
                await calendar.freebusy.query({

                    requestBody: {

                        timeMin:
                            new Date()
                                .toISOString(),

                        timeMax:
                            new Date(
                                Date.now() +
                                24 *
                                60 *
                                60 *
                                1000
                            ).toISOString(),

                        items: [
                            {
                                id:
                                    calendarId
                            }
                        ]

                    }

                });


            // ========================================================
            // RETURN RESULT
            // ========================================================

            res.json({

                success: true,

                message:
                    "Google Calendar access is working.",

                engineer:
                    engineer["Name"],

                calendarId:
                    calendarId,

                calendars:
                    response.data.calendars

            });


        } catch (error) {

            console.error(
                "Google Calendar test failed:"
            );

            console.error(
                error.response?.data ||
                error
            );


            res.status(500).json({

                success: false,

                message:
                    "Google Calendar access failed.",

                error:
                    error.response?.data ||
                    error.message

            });

        }

    }
);


// ============================================================
// UPDATE GOOGLE CALENDAR EVENT DURATION
// ============================================================

async function updateCalendarEventDuration(
    calendarId,
    eventId,
    appointmentDate,
    appointmentTime,
    appointmentDuration
) {

    // ============================================================
    // CREATE START TIME
    // ============================================================

    const startDateTime =
        `${appointmentDate}T${appointmentTime}:00+08:00`;


    const startTime =
        new Date(startDateTime);


    // ============================================================
    // VALIDATE START TIME
    // ============================================================

    if (
        isNaN(
            startTime.getTime()
        )
    ) {

        throw new Error(
            `Invalid appointment date/time: ${appointmentDate} ${appointmentTime}`
        );

    }


    // ============================================================
    // CALCULATE NEW END TIME
    // ============================================================

    const endTime =
        new Date(
            startTime.getTime() +
            Number(
                appointmentDuration
            ) *
            60 *
            60 *
            1000
        );


    console.log(
        "Calendar start:",
        startTime.toISOString()
    );

    console.log(
        "Calendar end:",
        endTime.toISOString()
    );


    // ============================================================
    // UPDATE EXISTING EVENT
    // ============================================================

    const response =
        await calendar.events.patch({

            calendarId:
                calendarId,

            eventId:
                eventId,

            requestBody: {

                start: {

                    dateTime:
                        startTime.toISOString(),

                    timeZone:
                        "Asia/Singapore"

                },

                end: {

                    dateTime:
                        endTime.toISOString(),

                    timeZone:
                        "Asia/Singapore"

                }

            }

        });


    return response.data;

}


// RESCHEDULE APPOINTMENT

app.post(
    "/api/tickets/:ticketId/reschedule",
    async (req, res) => {

        try {

            const ticketId =
                req.params.ticketId;


            const {
                appointmentDate,
                appointmentTime,
                appointmentDuration,
                onsiteWorkType
            } = req.body;


            // ========================================================
            // VALIDATE DATE
            // ========================================================

            if (!appointmentDate) {

                return res.status(400).json({
                    success: false,
                    message:
                        "New appointment date is required."
                });

            }


            // ========================================================
            // VALIDATE TIME
            // ========================================================

            if (!appointmentTime) {

                return res.status(400).json({
                    success: false,
                    message:
                        "New appointment time is required."
                });

            }


            // ========================================================
            // READ WORKBOOK
            // ========================================================

            const workbook =
                XLSX.readFile(excelFile);


            const ticketSheet =
                workbook.Sheets["Ticket"];


            if (!ticketSheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Ticket sheet not found."
                });

            }


            const tickets =
                XLSX.utils.sheet_to_json(
                    ticketSheet
                );


            // ========================================================
            // FIND TICKET
            // ========================================================

            const ticketIndex =
                tickets.findIndex(
                    ticket =>
                        String(
                            ticket["Ticket ID"]
                        ).trim() ===
                        String(ticketId).trim()
                );


            if (ticketIndex === -1) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Ticket not found."
                });

            }


            const ticket =
                tickets[ticketIndex];

          
            // ========================================================
            // CHECK APPOINTMENT STATUS
            // ========================================================

            const appointmentStatus =
                String(
                    ticket["Appointment Status"] || ""
                ).trim();


            if (
                appointmentStatus !==
                "Reschedule Required"
            ) {

                return res.status(400).json({
                    success: false,
                    message:
                        "This appointment is not currently marked as Reschedule Required."
                });

            }


            // ========================================================
            // CHECK SUPPORT TYPE
            // ========================================================

            const supportType =
                String(
                    ticket["Support Type"] || ""
                )
                    .trim()
                    .toLowerCase();


            const isRemoteSupport =
                supportType ===
                "remote support";


            const isOnSiteSupport =
                supportType ===
                    "on-site support" ||
                supportType ===
                    "on site support";


            if (
                !isRemoteSupport &&
                !isOnSiteSupport
            ) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Only Remote Support and On-site Support appointments can be rescheduled."
                });

            }


            // ========================================================
            // GET ON-SITE WORK TYPE
            // ========================================================

            let effectiveOnsiteWorkType =
                String(
                    onsiteWorkType ||
                    ticket["On-site Work Type"] ||
                    ""
                ).trim();


            // Remote Support does not use an on-site work type

            if (isRemoteSupport) {

                effectiveOnsiteWorkType = "";

            }


            // ========================================================
            // VALIDATE ON-SITE WORK TYPE
            // ========================================================

            if (isOnSiteSupport) {

                if (
                    !effectiveOnsiteWorkType
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "On-site Work Type is required for On-site Support."
                    });

                }


                if (
                    effectiveOnsiteWorkType !== "Ad Hoc" &&
                    effectiveOnsiteWorkType !== "Maintenance" &&
                    effectiveOnsiteWorkType !== "Project"
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Invalid On-site Work Type."
                    });

                }

            }


            // ========================================================
            // GET ENTERED DURATION
            //
            // For:
            // Ad Hoc       = hours
            // Maintenance = hours
            // Project     = days
            // ========================================================

            const enteredDuration =
                Number(
                    appointmentDuration !== undefined &&
                    appointmentDuration !== ""
                        ? appointmentDuration
                        : (
                            effectiveOnsiteWorkType === "Project"
                                ? Number(
                                    ticket["Appointment Duration"] || 0
                                ) / 8
                                : Number(
                                    ticket["Appointment Duration"] || 0
                                )
                        )
                );


            if (
                !Number.isFinite(
                    enteredDuration
                ) ||
                enteredDuration <= 0
            ) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Appointment duration must be greater than 0."
                });

            }


            // ========================================================
            // FINAL DURATION
            //
            // This is ALWAYS stored/sent to Google Calendar in HOURS.
            // ========================================================

            let duration =
                enteredDuration;


            // ========================================================
            // AD HOC
            // Minimum 2 hours
            // 0.5 hour increments
            // ========================================================

            if (
                isOnSiteSupport &&
                effectiveOnsiteWorkType === "Ad Hoc"
            ) {

                if (
                    enteredDuration < 2 ||
                    enteredDuration % 0.5 !== 0
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Ad Hoc work requires at least 2 hours and must use 0.5 hour increments."
                    });

                }

                duration =
                    enteredDuration;

            }


            // ========================================================
            // MAINTENANCE
            // Minimum 1 hour
            // 0.5 hour increments
            // ========================================================

            else if (
                isOnSiteSupport &&
                effectiveOnsiteWorkType === "Maintenance"
            ) {

                if (
                    enteredDuration < 1 ||
                    enteredDuration % 0.5 !== 0
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Maintenance work requires at least 1 hour and must use 0.5 hour increments."
                    });

                }

                duration =
                    enteredDuration;

            }


            // ========================================================
            // PROJECT
            // Minimum 1 whole day
            //
            // 1 day = 8 hours
            // ========================================================

            else if (
                isOnSiteSupport &&
                effectiveOnsiteWorkType === "Project"
            ) {

                if (
                    !Number.isInteger(
                        enteredDuration
                    ) ||
                    enteredDuration < 1
                ) {

                    return res.status(400).json({
                        success: false,
                        message:
                            "Project duration must be at least 1 whole day."
                    });

                }


                duration =
                    enteredDuration * 8;

            }


            // ========================================================
            // FINAL DURATION SAFETY CHECK
            // ========================================================

            if (
                !Number.isFinite(duration) ||
                duration <= 0
            ) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Invalid appointment duration."
                });

            }


            // ========================================================
            // GET ASSIGNED ENGINEER
            // ========================================================

            const engineerName =
                String(
                    ticket["Assigned Engineer"] || ""
                ).trim();


            if (!engineerName) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Please assign an engineer before rescheduling the appointment."
                });

            }


            // ========================================================
            // READ STAFF SHEET
            // ========================================================

            const staffSheet =
                workbook.Sheets["Staff"];


            if (!staffSheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Staff sheet not found."
                });

            }


            const staff =
                XLSX.utils.sheet_to_json(
                    staffSheet
                );


            // ========================================================
            // FIND ENGINEER
            // ========================================================

            const engineer =
                staff.find(
                    person =>
                        String(
                            person["Name"] || ""
                        )
                            .trim()
                            .toLowerCase() ===
                        engineerName.toLowerCase()
                );


            if (!engineer) {

                return res.status(404).json({
                    success: false,
                    message:
                        `Engineer "${engineerName}" could not be found in the Staff sheet.`
                });

            }


            // ========================================================
            // GET CALENDAR ID
            // ========================================================

            const calendarId =
                String(
                    engineer["Calendar ID"] || ""
                ).trim();


            if (!calendarId) {

                return res.status(400).json({
                    success: false,
                    message:
                        `Engineer "${engineerName}" does not have a Calendar ID.`
                });

            }


            // ========================================================
            // EXISTING GOOGLE CALENDAR EVENT
            // ========================================================

            const calendarEventId =
                String(
                    ticket["Calendar Event ID"] || ""
                ).trim();


            // ========================================================
            // CHECK CREDITS BEFORE MODIFYING CALENDAR
            //
            // IMPORTANT:
            // duration is already in HOURS here.
            // ========================================================

            const oldCreditsCharged =
                Number(
                    ticket["Credits Charged"] || 0
                );


            const creditDifference =
                duration -
                oldCreditsCharged;


            if (
                creditDifference > 0
            ) {

                const customerSheet =
                    workbook.Sheets["Customer"];


                if (!customerSheet) {

                    return res.status(500).json({
                        success: false,
                        message:
                            "Customer sheet not found."
                    });

                }


                const customers =
                    XLSX.utils.sheet_to_json(
                        customerSheet
                    );


                const customer =
                    customers.find(
                        item =>
                            String(
                                item["Customer ID"]
                            ).trim() ===
                            String(
                                ticket["Customer ID"]
                            ).trim()
                    );


                if (!customer) {

                    return res.status(404).json({
                        success: false,
                        message:
                            "Customer for this ticket could not be found."
                    });

                }


                const currentCredits =
                    Number(
                        customer["Credits"] || 0
                    );


                if (
                    currentCredits <
                    creditDifference
                ) {

                    return res.status(400).json({

                        success: false,

                        message:
                            `Customer does not have enough credits for the new duration. Required: ${creditDifference}, Available: ${currentCredits}.`,

                        requiredCredits:
                            creditDifference,

                        availableCredits:
                            currentCredits

                    });

                }

            }


            // ========================================================
            // CHECK GOOGLE CALENDAR AVAILABILITY
            //
            // duration is in HOURS
            // ========================================================

            console.log(
                "Checking calendar availability for reschedule..."
            );


            const availability =
                await checkCalendarAvailability(
                    calendarId,
                    appointmentDate,
                    appointmentTime,
                    duration,
                    calendarEventId
                );


            if (
                !availability.available
            ) {

                return res.status(409).json({

                    success: false,

                    message:
                        "The assigned engineer is not available at the selected appointment time.",

                    busyPeriods:
                        availability.busyPeriods || []

                });

            }


            // ========================================================
            // UPDATE GOOGLE CALENDAR EVENT
            // ========================================================

            let finalCalendarEventId =
                calendarEventId;


            if (calendarEventId) {

                console.log(
                    "Updating existing Google Calendar event..."
                );


                await updateCalendarEventDuration(
                    calendarId,
                    calendarEventId,
                    appointmentDate,
                    appointmentTime,
                    duration
                );

            } else {

                console.log(
                    "No existing Calendar Event ID. Creating new event..."
                );


                const googleEvent =
                    await createCalendarEvent(
                        calendarId,
                        ticketId,
                        ticket["Customer Name"],
                        ticket["Issue"],
                        engineer["Name"],
                        appointmentDate,
                        appointmentTime,
                        duration
                    );


                finalCalendarEventId =
                    googleEvent.id;

            }


            // ========================================================
            // UPDATE CUSTOMER CREDITS
            // ========================================================

            if (
                creditDifference !== 0
            ) {

                const customerSheet =
                    workbook.Sheets["Customer"];


                const customers =
                    XLSX.utils.sheet_to_json(
                        customerSheet
                    );


                const customerIndex =
                    customers.findIndex(
                        customer =>
                            String(
                                customer["Customer ID"]
                            ).trim() ===
                            String(
                                ticket["Customer ID"]
                            ).trim()
                    );


                if (
                    customerIndex === -1
                ) {

                    return res.status(404).json({
                        success: false,
                        message:
                            "Customer for this ticket could not be found."
                    });

                }


                const customer =
                    customers[customerIndex];


                const currentCredits =
                    Number(
                        customer["Credits"] || 0
                    );


                customer["Credits"] =
                    currentCredits -
                    creditDifference;


                customers[customerIndex] =
                    customer;


                workbook.Sheets["Customer"] =
                    XLSX.utils.json_to_sheet(
                        customers
                    );

            }


            // ========================================================
            // UPDATE TICKET
            // ========================================================

            ticket["Appointment Date"] =
                appointmentDate;


            ticket["Appointment Time"] =
                appointmentTime;


            // IMPORTANT:
            // Excel stores Project duration as HOURS.
            // Example: 2 Project days = 16 hours.

            ticket["Appointment Duration"] =
                duration;


            ticket["On-site Work Type"] =
                effectiveOnsiteWorkType;


            ticket["Appointment Status"] =
                "Confirmed";


            ticket["Calendar Event ID"] =
                finalCalendarEventId;


            ticket["Credits Charged"] =
                duration;


            tickets[ticketIndex] =
                ticket;


            workbook.Sheets["Ticket"] =
                XLSX.utils.json_to_sheet(
                    tickets
                );


            // ========================================================
            // SAVE EXCEL
            // ========================================================

            XLSX.writeFile(
                workbook,
                excelFile
            );


            // ========================================================
            // SUCCESS
            // ========================================================

            return res.json({

                success: true,

                message:
                    "Appointment rescheduled successfully.",

                ticketId:
                    ticketId,

                appointmentDate:
                    appointmentDate,

                appointmentTime:
                    appointmentTime,

                appointmentDuration:
                    duration,

                onsiteWorkType:
                    effectiveOnsiteWorkType,

                appointmentStatus:
                    "Confirmed",

                calendarEventId:
                    finalCalendarEventId

            });


        } catch (error) {

            console.error(
                "Error rescheduling appointment:"
            );

            console.error(error);


            return res.status(500).json({

                success: false,

                message:
                    "Could not reschedule appointment.",

                error:
                    error.response?.data ||
                    error.message

            });

        }

    }
);

// CHECK ENGINEER AVAILABILITY
app.get(
    "/api/appointments/availability",
    async (req, res) => {

        try {

            const {
                engineerName,
                appointmentDate,
                appointmentTime,
                appointmentDuration,
                onsiteWorkType
            } = req.query;


            // ========================================================
            // VALIDATION
            // ========================================================

            if (!engineerName) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Engineer is required."
                });

            }


            if (!appointmentDate) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Appointment date is required."
                });

            }


            if (!appointmentTime) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Appointment time is required."
                });

            }

            // CONVERT DURATION TO HOURS
                let duration = Number(appointmentDuration) || 1;
            // READ STAFF
 

            const workbook =
                XLSX.readFile(excelFile);


            const staffSheet =
                workbook.Sheets["Staff"];


            if (!staffSheet) {

                return res.status(500).json({
                    success: false,
                    message:
                        "Staff sheet not found."
                });

            }


            const staff =
                XLSX.utils.sheet_to_json(
                    staffSheet
                );


            // ========================================================
            // FIND ENGINEER
            // ========================================================

            const engineer =
                staff.find(
                    person =>
                        String(
                            person["Name"] ||
                            ""
                        )
                            .trim()
                            .toLowerCase() ===
                        String(
                            engineerName
                        )
                            .trim()
                            .toLowerCase()
                );


            if (!engineer) {

                return res.status(404).json({
                    success: false,
                    message:
                        "Engineer could not be found."
                });

            }


            // ========================================================
            // GET CALENDAR ID
            // ========================================================

            const calendarId =
                String(
                    engineer["Calendar ID"] ||
                    ""
                ).trim();


            if (!calendarId) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Engineer does not have a Calendar ID."
                });

            }


            // ========================================================
            // CHECK CALENDAR
            // ========================================================

            const result =
                await checkCalendarAvailability(
                    calendarId,
                    appointmentDate,
                    appointmentTime,
                    duration
                );


            // ========================================================
            // RESPONSE
            // ========================================================

            res.json({

                success: true,

                engineer:
                    engineer["Name"],

                calendarId:
                    calendarId,

                appointmentDate:
                    appointmentDate,

                appointmentTime:
                    appointmentTime,

                onsiteWorkType:
                    onsiteWorkType || "",

                appointmentDuration:
                    duration,

                available:
                    result.available,

                busyPeriods:
                    result.busyPeriods

            });


        } catch (error) {

            console.error(
                "Calendar availability check failed:"
            );

            console.error(error);


            res.status(500).json({

                success: false,

                message:
                    "Failed to check calendar availability.",

                error:
                    error.response?.data ||
                    error.message

            });

        }

    }
);
// ============================================================
// CREATE GOOGLE CALENDAR EVENT
// ============================================================

async function createCalendarEvent(
    calendarId,
    ticketId,
    customerName,
    issue,
    engineerName,
    appointmentDate,
    appointmentTime,
    appointmentDuration,
    onsiteWorkType = ""
) {

    try {

        // ========================================================
        // APPOINTMENT DURATION IS ALREADY IN HOURS
        // ========================================================
        //
        // Project conversion happens BEFORE this function.
        //
        // Example:
        // Project 2 days
        // -> reschedule route converts to 16 hours
        // -> this function receives 16
        //
        // DO NOT multiply by 8 here.
        // ========================================================

        const durationHours =
            Number(appointmentDuration);


        if (
            !Number.isFinite(durationHours) ||
            durationHours <= 0
        ) {

            throw new Error(
                "Invalid appointment duration."
            );

        }


        // ========================================================
        // SINGAPORE TIMEZONE
        // ========================================================

        const startDateTime =
            `${appointmentDate}T${appointmentTime}:00+08:00`;


        const startTime =
            new Date(startDateTime);


        if (
            isNaN(
                startTime.getTime()
            )
        ) {

            throw new Error(
                `Invalid appointment date/time: ${appointmentDate} ${appointmentTime}`
            );

        }


        // ========================================================
        // CALCULATE END TIME
        // ========================================================

        const endTime =
            new Date(
                startTime.getTime() +
                durationHours *
                60 *
                60 *
                1000
            );


        // ========================================================
        // CREATE EVENT
        // ========================================================

        const event = {

            summary:
                `IT Support - ${ticketId}`,

            description:
                `Customer: ${customerName}\n` +
                `Issue: ${issue}\n` +
                `Engineer: ${engineerName}\n` +
                `Ticket ID: ${ticketId}\n` +
                `Work Type: ${onsiteWorkType || "Remote Support"}\n` +
                `Duration: ${durationHours} hours`,

            start: {

                dateTime:
                    startTime.toISOString(),

                timeZone:
                    "Asia/Singapore"

            },

            end: {

                dateTime:
                    endTime.toISOString(),

                timeZone:
                    "Asia/Singapore"

            }

        };


        // ========================================================
        // INSERT GOOGLE CALENDAR EVENT
        // ========================================================

        const response =
            await calendar.events.insert({

                calendarId:
                    calendarId,

                requestBody:
                    event

            });


        console.log(
            `Google Calendar event created: ${response.data.id}`
        );


        return response.data;


    } catch (error) {

        console.error(
            "Google Calendar event creation failed:"
        );

        console.error(
            error.response?.data ||
            error.message
        );

        throw error;

    }

}

// EMAIL SERVICE REPORT


app.post("/api/tickets/:ticketId/email-report", async (req, res) => {

    try {

        const ticketId = req.params.ticketId;
        const pdfBase64 = req.body.pdfBase64;

        if (!pdfBase64) {
            return res.status(400).json({
                success: false,
                message: "PDF data is required."
            });
        }

        const tickets = getSheetData("Ticket");

        const ticket = tickets.find(
            item =>
                String(item["Ticket ID"]).trim() ===
                String(ticketId).trim()
        );

        if (!ticket) {
            return res.status(404).json({
                success: false,
                message: "Ticket not found."
            });
        }

        const customerEmail =
            String(ticket["Email"] || "").trim();

        if (!customerEmail) {
            return res.status(400).json({
                success: false,
                message: "This ticket does not have a customer email address."
            });
        }

        const customerName =
            ticket["Customer Name"] ||
            "Customer";

        const supportType =
            ticket["Support Type"] ||
            "Not recorded";

        const engineer =
            ticket["Assigned Engineer"] ||
            "Not assigned";

        const issue =
            ticket["Issue"] ||
            "No issue recorded";

        const serviceResult =
            ticket["Service Result"] ||
            "No service result recorded";

        const normalizedSupportType =
            String(supportType)
                .trim()
                .toLowerCase();

        const hasAppointment =
            normalizedSupportType === "remote support" ||
            normalizedSupportType === "on-site support" ||
            normalizedSupportType === "on site support";

        let emailHtml = `

            <div style="
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333;
            ">

                <h2>IT Support Service Report</h2>

                <p>
                    Dear ${customerName},
                </p>

                <p>
                    Your IT support service report is now available.
                    Please find the service report attached to this
                    email for your records.
                </p>

                <hr>

                <h3>Ticket Information</h3>

                <p>
                    <strong>Ticket ID:</strong>
                    ${ticketId}
                </p>

                <p>
                    <strong>Support Type:</strong>
                    ${supportType}
                </p>

                <p>
                    <strong>Assigned Engineer:</strong>
                    ${engineer}
                </p>

                <p>
                    <strong>Issue:</strong><br>
                    ${issue}
                </p>

                <p>
                    <strong>Service Result:</strong><br>
                    ${serviceResult}
                </p>
        `;

        if (hasAppointment) {

            emailHtml += `

                <hr>

                <h3>Appointment Information</h3>

                <p>
                    <strong>Appointment Date:</strong>
                    ${ticket["Appointment Date"] || "Not recorded"}
                </p>

                <p>
                    <strong>Appointment Time:</strong>
                    ${ticket["Appointment Time"] || "Not recorded"}
                </p>

                <p>
                    <strong>Duration:</strong>
                    ${
                        ticket["Appointment Duration"]
                            ? ticket["Appointment Duration"] + " hours"
                            : "Not recorded"
                    }
                </p>

                <p>
                    <strong>Appointment Status:</strong>
                    ${ticket["Appointment Status"] || "Not recorded"}
                </p>

                <hr>

                <h3>Employee Work Session</h3>

                <p>
                    <strong>Employee Sign In:</strong>
                    ${ticket["Employee Sign In"] || "Not recorded"}
                </p>

                <p>
                    <strong>Employee Sign Out:</strong>
                    ${ticket["Employee Sign Out"] || "Not recorded"}
                </p>

                <p>
                    <strong>Work Duration:</strong>
                    ${
                        ticket["Work Duration"]
                            ? ticket["Work Duration"] + " hours"
                            : "Not recorded"
                    }
                </p>

            `;
        }

        emailHtml += `

                <hr>

                <p>
                    Thank you for using our IT Support service.
                </p>

                <p>
                    This is an automated email.
                </p>

            </div>
        `;

        // --------------------------------------------------------
        // CONVERT BASE64 PDF TO BUFFER
        // --------------------------------------------------------

        const cleanBase64 =
            pdfBase64.includes(",")
                ? pdfBase64.split(",")[1]
                : pdfBase64;

        const pdfBuffer =
            Buffer.from(
                cleanBase64,
                "base64"
            );

        // --------------------------------------------------------
        // SEND EMAIL USING RESEND
        // --------------------------------------------------------

        const {
            data,
            error
        } = await resend.emails.send({

            from:
                "IT Support <onboarding@resend.dev>",

            to:
                [customerEmail],

            subject:
                `IT Support Service Report - ${ticketId}`,

            html:
                emailHtml,

            attachments: [

                {
                    filename:
                        `Service-Report-${ticketId}.pdf`,

                    content:
                        pdfBuffer
                }

            ]

        });

        if (error) {

            console.error(
                "Resend error:",
                error
            );

            return res.status(500).json({

                success: false,

                message:
                    "Resend failed to send the email.",

                error:
                    error.message ||
                    String(error)

            });

        }

        console.log("======================================");
        console.log("SERVICE REPORT EMAIL SENT");
        console.log("Ticket:", ticketId);
        console.log("Customer:", customerName);
        console.log("Email:", customerEmail);
        console.log("Resend ID:", data?.id);
        console.log("======================================");

        return res.json({

            success: true,

            message:
                "Service report emailed successfully.",

            email:
                customerEmail,

            emailId:
                data?.id || ""

        });

    } catch (error) {

        console.error(
            "Error sending service report email:"
        );

        console.error(error);

        return res.status(500).json({

            success: false,

            message:
                "Could not send service report email.",

            error:
                error.message

        });

    }

});
app.post("/api/staff/login", (req, res) => {

    try {

        const { staffId, password } = req.body;

        const staff = getSheetData("Staff");

        const foundStaff = staff.find(row =>
            String(row["Staff ID"]).trim() === String(staffId).trim() &&
            String(row["Password"]).trim() === String(password).trim()
        );

        if (!foundStaff) {

            return res.status(401).json({
                success: false,
                message: "Invalid Staff ID or password."
            });

        }

        res.json({
            success: true,
            message: "Login successful.",
            staff: {
                staffId: foundStaff["Staff ID"],
                name: foundStaff["Name"],
                email: foundStaff["Email"],
                role: foundStaff["Role"],
                department: foundStaff["Department"],
                calendarId: foundStaff["Calendar ID"]
            }
        });

    } catch (error) {

        console.error("Staff login error:", error);

        res.status(500).json({
            success: false,
            message: "Server error during staff login."
        });

    }

});
app.post("/api/tickets/:ticketId/reschedule", async (req, res) => {

    try {

        const ticketId = req.params.ticketId;

        const {
    assignedEngineer,
    status,
    appointmentDuration,
    onsiteWorkType,
    appointmentStatus,
    serviceResult
} = req.body;


        // ============================================================
        // VALIDATE REQUIRED FIELDS
        // ============================================================

        if (
            !assignedEngineer ||
            !appointmentDate ||
            !appointmentTime ||
            !appointmentDuration
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Engineer, date, time and duration are required."
            });

        }


        // ============================================================
        // CONVERT DURATION
        // ============================================================

        const enteredDuration =
            Number(appointmentDuration);


        if (
            !enteredDuration ||
            enteredDuration <= 0
        ) {

            return res.status(400).json({
                success: false,
                message:
                    "Appointment duration must be greater than 0."
            });

        }


        // ------------------------------------------------------------
        // PROJECT = DAYS
        // 1 PROJECT DAY = 8 HOURS
        // ------------------------------------------------------------

        let duration =
            enteredDuration;


        if (onsiteWorkType === "Project") {

            if (!Number.isInteger(enteredDuration)) {

                return res.status(400).json({
                    success: false,
                    message:
                        "Project duration must be a whole number of days."
                });

            }

            duration =
                enteredDuration * 8;

        }


        // ============================================================
        // LOAD EXCEL
        // ============================================================

        const workbook =
            XLSX.readFile(excelFile);

        const ticketSheet =
            workbook.Sheets["Ticket"];

        const staffSheet =
            workbook.Sheets["Staff"];


        if (!ticketSheet) {

            return res.status(500).json({
                success: false,
                message:
                    "Ticket sheet not found."
            });

        }


        if (!staffSheet) {

            return res.status(500).json({
                success: false,
                message:
                    "Staff sheet not found."
            });

        }


        const tickets =
            XLSX.utils.sheet_to_json(
                ticketSheet
            );

        const staff =
            XLSX.utils.sheet_to_json(
                staffSheet
            );


        // ============================================================
        // FIND TICKET
        // ============================================================

        const ticketIndex =
            tickets.findIndex(
                ticket =>
                    String(
                        ticket["Ticket ID"]
                    ).trim() ===
                    String(
                        ticketId
                    ).trim()
            );


        if (ticketIndex === -1) {

            return res.status(404).json({
                success: false,
                message:
                    "Ticket not found."
            });

        }


        const ticket =
            tickets[ticketIndex];

       
        // ============================================================
        // FIND SELECTED ENGINEER
        // ============================================================

        const selectedStaff =
            staff.find(
                person =>
                    String(
                        person["Name"] || ""
                    )
                        .trim()
                        .toLowerCase() ===
                    String(
                        assignedEngineer
                    )
                        .trim()
                        .toLowerCase()
            );


        if (!selectedStaff) {

            return res.status(400).json({
                success: false,
                message:
                    "Selected engineer was not found in the Staff sheet."
            });

        }


        // ============================================================
        // GET NEW CALENDAR ID
        // ============================================================

        const newCalendarId =
            selectedStaff["Calendar ID"] ||
            selectedStaff["Email"];


        if (!newCalendarId) {

            return res.status(400).json({
                success: false,
                message:
                    `No Calendar ID or Email is configured for ${assignedEngineer}.`
            });

        }


        // ============================================================
        // CHECK SELECTED ENGINEER'S GOOGLE CALENDAR
        // ============================================================

        const availability =
            await checkCalendarAvailability(
                newCalendarId,
                appointmentDate,
                appointmentTime,
                duration,

                // Ignore the old event if
                // staying with the same engineer.
                String(
                    ticket["Assigned Engineer"] || ""
                )
                    .trim()
                    .toLowerCase() ===
                String(
                    assignedEngineer
                )
                    .trim()
                    .toLowerCase()
                    ?
                    ticket["Calendar Event ID"]
                    :
                    ""
            );


        if (!availability.available) {

            return res.status(409).json({

                success: false,

                available: false,

                message:
                    `${assignedEngineer} is not available at the selected appointment time.`,

                busyPeriods:
                    availability.busyPeriods

            });

        }


        // ============================================================
        // OLD ENGINEER INFORMATION
        // ============================================================

        const oldEngineerName =
            String(
                ticket["Assigned Engineer"] || ""
            ).trim();


        const oldStaff =
            staff.find(
                person =>
                    String(
                        person["Name"] || ""
                    )
                        .trim()
                        .toLowerCase() ===
                    oldEngineerName.toLowerCase()
            );


        const oldCalendarId =
            oldStaff?.["Calendar ID"] ||
            oldStaff?.["Email"] ||
            "";


        const oldEventId =
            ticket["Calendar Event ID"] ||
            "";


        const engineerChanged =
            oldEngineerName.toLowerCase() !==
            String(
                assignedEngineer
            )
                .trim()
                .toLowerCase();


        // ============================================================
        // CREATE / UPDATE GOOGLE CALENDAR EVENT
        // ============================================================

        let newEventId =
            oldEventId;


        if (
            !engineerChanged &&
            oldEventId
        ) {

            // --------------------------------------------------------
            // SAME ENGINEER
            // Update the existing event.
            //
            // duration is already in HOURS here.
            // --------------------------------------------------------

            await updateCalendarEventDuration(
                newCalendarId,
                oldEventId,
                appointmentDate,
                appointmentTime,
                duration
            );


        } else {

            // --------------------------------------------------------
            // ENGINEER CHANGED
            // OR OLD EVENT DOES NOT EXIST
            //
            // Create the new event first.
            // --------------------------------------------------------

            const newEvent =
                await createCalendarEvent(
                    newCalendarId,
                    ticketId,
                    ticket["Customer Name"],
                    ticket["Issue"],
                    assignedEngineer,
                    appointmentDate,
                    appointmentTime,
                    duration,
                    onsiteWorkType
                );


            newEventId =
                newEvent.id;


            // --------------------------------------------------------
            // DELETE OLD EVENT
            // ONLY AFTER NEW EVENT WAS SUCCESSFULLY CREATED
            // --------------------------------------------------------

            if (
                engineerChanged &&
                oldCalendarId &&
                oldEventId
            ) {

                await deleteCalendarEvent(
                    oldCalendarId,
                    oldEventId
                );

            }

        }


        // ============================================================
        // UPDATE EXCEL
        // ============================================================

        ticket["Assigned Engineer"] =
            assignedEngineer;

        ticket["Appointment Date"] =
            appointmentDate;

        ticket["Appointment Time"] =
            appointmentTime;


        // Store HOURS in Excel.
        //
        // Example:
        // Ad Hoc 2.5 hours  -> 2.5
        // Maintenance 1.5  -> 1.5
        // Project 2 days    -> 16
        //
        ticket["Appointment Duration"] =
            duration;


        ticket["On-site Work Type"] =
            onsiteWorkType || "";


        ticket["Appointment Status"] =
            "Confirmed";


        ticket["Calendar Event ID"] =
            newEventId;


        const newTicketSheet =
            XLSX.utils.json_to_sheet(
                tickets
            );


        workbook.Sheets["Ticket"] =
            newTicketSheet;


        XLSX.writeFile(
            workbook,
            excelFile
        );


        // ============================================================
        // RESPONSE
        // ============================================================

        return res.json({

            success: true,

            message:
                "Appointment successfully rescheduled.",

            assignedEngineer:
                assignedEngineer,

            appointmentDate:
                appointmentDate,

            appointmentTime:
                appointmentTime,

            appointmentDuration:
                duration,

            onsiteWorkType:
                onsiteWorkType || "",

            calendarId:
                newCalendarId,

            calendarEventId:
                newEventId

        });


    } catch (error) {

        console.error(
            "RESCHEDULE APPOINTMENT ERROR:",
            error
        );


        return res.status(500).json({

            success: false,

            message:
                error.message ||
                "Failed to reschedule appointment."

        });

    }

});
// GET ALL KNOWLEDGE BASE ARTICLES
app.get("/api/knowledge-base", (req, res) => {
    try {
        const workbook = XLSX.readFile(excelFile);

        const sheet = workbook.Sheets["Knowledge-Base"];

        if (!sheet) {
            return res.status(404).json({
                success: false,
                message: "Knowledge-Base sheet not found."
            });
        }

        const articles = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        res.json({
            success: true,
            articles
        });

    } catch (error) {
        console.error("Knowledge Base error:", error);

        res.status(500).json({
            success: false,
            message: "Failed to load Knowledge Base."
        });
    }
});

// START SERVER

app.listen(
    PORT,
    () => {

        console.log(
            `Server running on: http://localhost:${PORT}`
        );

    }
);