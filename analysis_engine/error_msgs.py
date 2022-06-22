import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s: %(levelname)s - %(message)s",
    datefmt="%d-%b-%y %H:%M:%S",
)
logger = logging.getLogger(__name__)


class ProjectNameError(Exception):
    pass


class ProjectGroupError(Exception):
    pass


class ProjectStageError(Exception):
    pass


def latest_project_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.critical(e + " has not been found in the project_info document.")
        raise ProjectNameError(
            "Project names in the latest master and project_info must match. Program stopping. Please amend."
        )


def historic_project_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.info(
            f"Project name {e} in master {error_cases[e]} not in project information "
            f"document. Please make sure project names are consistent or it could cause "
            f"analysis engine to crash or inaccurate analysis."
        )


def latest_group_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.critical(
                e + " does not have a recognised Group value in the project information document.")
        raise ProjectGroupError(
            "Project names in the latest master and project_info must match. Program stopping. Please amend."
        )


def historic_group_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.info(
                e + " does not have a recognised Group value in the project information document. "
                    "As not in current master, analysis_engine not stopping. But please this could "
                    "cause a crash or inaccurate analysis and should be corrected.")