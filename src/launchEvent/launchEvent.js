function onNewAppointment(event) {
    console.log(event)
}

function onAppointmentSend(event) {
    console.log(event)
}

if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onNewAppointment", onNewAppointment);
    Office.actions.associate("onAppointmentSend", onAppointmentSend);
}