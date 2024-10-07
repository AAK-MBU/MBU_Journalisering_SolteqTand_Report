"""This module contains the main process of the robot."""
import json
from datetime import datetime
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from itk_dev_shared_components.smtp import smtp_util
from robot_framework import config
from robot_framework.subprocesses.manuel_list_handler import ManuelListHandler


def process(orchestrator_connection: OrchestratorConnection) -> None:
    """Do the primary process of the robot."""
    orchestrator_connection.log_trace("Running process.")

    rpa_db_connection = orchestrator_connection.get_constant("DbConnectionString").value

    man_list = ManuelListHandler(rpa_db_connection)

    try:
        orchestrator_connection.log_trace("Sending e-mail.")
        oc_args_json = json.loads(orchestrator_connection.process_arguments)

        receiver = [oc_args_json['emailReceiver1'], oc_args_json['emailReceiver2']]
        sender = oc_args_json['emailSender']
        now = datetime.now().strftime("%d-%m-%y")
        subject = f"Journalisering - Solteq Tand - Manuel liste d. {now}"
        body = man_list.generate_list()

        smtp_util.send_email(
            receiver=receiver,
            sender=sender,
            subject=subject,
            body=body,
            html_body=True,
            smtp_server=config.SMTP_SERVER,
            smtp_port=config.SMTP_PORT
        )

        orchestrator_connection.log_trace(f"E-mail was sent to: {receiver}")

    except (Exception, RuntimeError) as e:
        raise e
