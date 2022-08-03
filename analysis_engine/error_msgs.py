import logging, sys

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


class ConfigurationError(Exception):
    pass


class InputError(Exception):
    pass


def config_issue():
    logger.critical(
        "Configuration file issue. Please check and make sure it's correct. Hint: It could be "
        "a file name(s) does not match what is in the config or is not in the correct folder."
    )
    sys.exit(1)


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


def historic_stage_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.info(
                f"Project name {e} in master {error_cases[e]} does not have a recognised stage name "
                f". Please make sure stage data is consistent with terminology in config file or it could cause "
                f"analysis engine to crash or inaccurate analysis."
            )


def latest_group_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.critical(
                e
                + " does not have a recognised group value in the project information document."
            )
        raise ProjectGroupError("Program stopping. Please amend.")


def latest_stage_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.critical(
                e + " does not have a recognised stage value in the latest master."
            )
        raise ProjectGroupError("Program stopping. Please amend.")


def historic_group_names_error(error_cases):
    if error_cases:
        for e in error_cases:
            logger.info(
                e
                + " does not have a recognised Group value in the project information document. "
                "As not in current master, analysis_engine not stopping. But please this could "
                "cause a crash or inaccurate analysis and should be corrected."
            )


def abbreviation_error(error_cases):
    if error_cases:
        for p in error_cases:
            logger.critical("No abbreviation provided for " + p + ".")
        raise ProjectNameError(
            "Abbreviations must be provided for all projects in project_info. Program stopping. Please amend"
        )


def not_recognised_project_group_or_stage(error_case):
    if error_case:
        for p in error_case:
            logger.critical(p + " not a recognised project or group or stage name")
        raise ProjectNameError(
            "Program stopping. Please check project, group or stage names being entered."
        )


def not_recognised_quarter(error_case):
    logger.critical(error_case + " not a recognised quarter.")
    raise InputError(
        "Program stopping. Please check project, group or stage names being entered."
    )


def not_recognised_quarter(error):
    raise InputError(
        error + " not a recognised quarter. Program stopped. Please re-enter."
    )


def no_query_keys():
    raise InputError(
        "Program stopping. You must enter a key name(s) using either --koi or --koi_fn."
    )


def get_error_list(seq: list):
    seen = set()
    seen_add = seen.add
    return [x for x in seq if not (x in seen or seen_add(x))]


def resourcing_keys(p_name, key_name):
    logger.info(f"{p_name} has reported text for the key {key_name}. It has been skipped. "
                f"This data needs to be changed to a number.")
