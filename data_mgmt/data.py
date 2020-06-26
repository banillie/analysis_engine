class MilestoneData:
    def __init__(self, master_data, baseline_index, data_to_return):
        self.master_data = master_data
        self.baseline_index = baseline_index
        self.data_to_return = data_to_return
        self.project_dict = {}

    def project_data(self, project_names):  # renamed to project_data
        """
        Given a list of project names in project_names,
        returns a dictionary containing data for each project.
        """
        # Provided a description of what method does, including the
        # expected parameters and return type. NB, we use " quotes, not '
        # quotes for docstrings like this.

        upper_dict = {}

        for name in project_names:
            lower_dict = {}
            raw_list = []
            try:
                p_data = self.master_data[
                    self.baseline_index[name][self.data_to_return]
                ].data[name]
                for i in range(1, 50):
                    try:
                        try:
                            t = (
                                p_data["Approval MM" + str(i)],
                                p_data["Approval MM" + str(i) + " Forecast / Actual"],
                                p_data["Approval MM" + str(i) + " Notes"],
                            )
                            raw_list.append(t)
                        except KeyError:
                            t = (
                                p_data["Approval MM" + str(i)],
                                p_data["Approval MM" + str(i) + " Forecast - Actual"],
                                p_data["Approval MM" + str(i) + " Notes"],
                            )
                            raw_list.append(t)

                        t = (
                            p_data["Assurance MM" + str(i)],
                            p_data["Assurance MM" + str(i) + " Forecast - Actual"],
                            p_data["Assurance MM" + str(i) + " Notes"],
                        )
                        raw_list.append(t)

                    except KeyError:
                        pass

                for i in range(18, 67):
                    try:
                        t = (
                            p_data["Project MM" + str(i)],
                            p_data["Project MM" + str(i) + " Forecast - Actual"],
                            p_data["Project MM" + str(i) + " Notes"],
                        )
                        raw_list.append(t)
                    except KeyError:
                        pass
            except (KeyError, TypeError):
                print("yes")
                pass

            # put the list in chronological order
            sorted_list = sorted(raw_list, key=lambda k: (k[1] is None, k[1]))
            print(sorted_list)

            # loop to stop key names being the same. Not ideal as doesn't handle keys that may already have numbers as
            # strings at end of names. But still useful.
            for x in sorted_list:
                if x[0] is not None:
                    if x[0] in lower_dict:
                        for i in range(2, 15):
                            key_name = x[0] + " " + str(i)
                            if key_name in lower_dict:
                                continue
                            else:
                                lower_dict[key_name] = {x[1]: x[2]}
                                break
                    else:
                        lower_dict[x[0]] = {x[1]: x[2]}
                else:
                    pass

            upper_dict[name] = lower_dict

        self.project_dict = upper_dict

        return self.project_dict
