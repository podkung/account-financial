odoo.define("dynamic_accounts_report.financial_reports", function (require) {
    "use strict";
    var AbstractAction = require("web.AbstractAction");
    var core = require("web.core");
    var rpc = require("web.rpc");
    var QWeb = core.qweb;

    var ProfitAndLoss = AbstractAction.extend({
        template: "dfr_template_new",
        events: {
            "click #apply_filter": "apply_filter",
            "click #pdf": "print_pdf",
            "click #xlsx": "print_xlsx",
            "click .show-gl": "show_gl",
            "click .show-account": "show_account",
            "click .hide-account": "hide_account",
        },

        init: function (parent, action) {
            this._super(parent, action);
            this.currency = action.currency;
            this.report_lines = action.report_lines;
            this.wizard_id = action.context.wizard | null;
        },

        start: function () {
            var self = this;
            self.initial_render = true;
            rpc.query({
                model: "dynamic.balance.sheet.report",
                method: "create",
                args: [{}],
            }).then(function (t_res) {
                self.wizard_id = t_res;
                self.load_data(self.initial_render);
            });
        },

        load_data: function (initial_render = true) {
            var self = this;
            var action_title = self._title;
            try {
                self._rpc({
                    model: "dynamic.balance.sheet.report",
                    method: "view_report",
                    args: [[self.wizard_id], action_title],
                }).then(function (datas) {
                    if (initial_render) {
                        self.$(".filter_view_dfr").html(
                            QWeb.render("DfrFilterView", {
                                filter_data: datas.filters,
                                title: datas.name,
                            })
                        );
                        self.$el.find(".journals").select2({
                            placeholder: "Journals...",
                        });
                        self.$el.find(".account").select2({
                            placeholder: "Accounts...",
                        });
                        self.$el.find(".account-tag").select2({
                            placeholder: "Account Tag...",
                        });
                        self.$el.find(".analytics").select2({
                            placeholder: "Analytic Accounts...",
                        });
                        self.$el.find(".analytic-tag").select2({
                            placeholder: "Analytic Tag...",
                        });
                        self.$el.find(".target_move").select2({
                            placeholder: "Target Move...",
                        });
                        self.$el.find(".debit_credit").select2({
                            placeholder: "Debit/Credit...",
                        });
                        self.$el.find(".operating-units").select2({
                            placeholder: "Operating Units...",
                        });
                    }
                    self.$(".table_view_dfr").html(
                        QWeb.render("dfr_table", {
                            report_lines: datas.report_lines,
                            filter: datas.filters,
                            currency: datas.currency,
                            credit_total: datas.credit_total,
                            debit_total: datas.debit_total,
                            debit_balance: datas.debit_balance,
                            bs_lines: datas.bs_lines,
                        })
                    );
                });
            } catch (el) {
                console.log(el);
            }
        },

        show_gl: function (e) {
            var self = this;
            var account_id = $(e.target).attr("data-account-id");
            var action = {
                type: "ir.actions.client",
                name: "GL View",
                tag: "g_l",
                target: "new",
                context: {
                    account_id: account_id,
                    wizard_id: self.wizard_id,
                },
            };
            return this.do_action(action);
        },

        show_account: function (e) {
            $("tr[data-p-id=" + $(e.target).data("r-id") + "]").show();
            $(e.target)
                .removeClass("fa-caret-right show-account")
                .addClass("fa-caret-down hide-account");
        },

        hide_account: function (e) {
            $("tr[data-p-id=" + $(e.target).data("r-id") + "]").hide();
            $(e.target)
                .removeClass("fa-caret-down hide-account")
                .addClass("fa-caret-right show-account");
        },

        remove_bs_lines: function (data) {
            var r_ids = [];
            $(".show-account").each(function () {
                r_ids.push($(this).data("r-id"));
            });
            data.bs_lines = data.bs_lines.filter((line) => {
                return !r_ids.includes(line.p_id);
            });
            return data;
        },

        print_pdf: function (e) {
            e.preventDefault();
            var self = this;
            var action_title = self._title;
            self._rpc({
                model: "dynamic.balance.sheet.report",
                method: "view_report",
                args: [[self.wizard_id], action_title],
            }).then(function (data) {
                var action = {
                    type: "ir.actions.report",
                    report_type: "qweb-pdf",
                    report_name: "dynamic_accounts_report.balance_sheet",
                    report_file: "dynamic_accounts_report.balance_sheet",
                    data: {
                        report_data: self.remove_bs_lines(data),
                        report_name: action_title,
                    },
                    context: {
                        active_model: "dynamic.balance.sheet.report",
                        landscape: 1,
                        bs_report: true,
                    },
                    display_name: action_title,
                };
                return self.do_action(action);
            });
        },

        print_xlsx: function () {
            var self = this;
            var action_title = self._title;
            self._rpc({
                model: "dynamic.balance.sheet.report",
                method: "view_report",
                args: [[self.wizard_id], action_title],
            }).then(function (data) {
                var action = {
                    type: "ir_actions_dynamic_xlsx_download",
                    data: {
                        model: "dynamic.balance.sheet.report",
                        options: JSON.stringify(data.filters),
                        output_format: "xlsx",
                        report_data: action_title,
                        report_name: action_title,
                        dfr_data: JSON.stringify(self.remove_bs_lines(data).bs_lines),
                    },
                };
                return self.do_action(action);
            });
        },

        apply_filter: function (event) {
            event.preventDefault();
            var self = this;
            var i = 0;
            self.initial_render = false;
            var filter_data_selected = {};
            // Accounts
            var account_ids = [];
            var account_text = [];
            var account_res = document.getElementById("acc_res");
            var account_list = $(".account").select2("data");
            for (i = 0; i < account_list.length; i++) {
                if (account_list[i].element[0].selected === true) {
                    account_ids.push(parseInt(account_list[i].id, 10));
                    if (account_text.includes(account_list[i].text) === false) {
                        account_text.push(account_list[i].text);
                    }
                    account_res.value = account_text;
                    account_res.innerHTML = account_res.value;
                }
            }
            if (account_list.length === 0) {
                account_res.value = "";
                account_res.innerHTML = "";
            }
            filter_data_selected.account_ids = account_ids;

            // Journals
            var journal_ids = [];
            var journal_text = [];
            var journal_res = document.getElementById("journal_res");
            var journal_list = $(".journals").select2("data");
            for (i = 0; i < journal_list.length; i++) {
                if (journal_list[i].element[0].selected === true) {
                    journal_ids.push(parseInt(journal_list[i].id, 10));
                    if (journal_text.includes(journal_list[i].text) === false) {
                        journal_text.push(journal_list[i].text);
                    }
                    journal_res.value = journal_text;
                    journal_res.innerHTML = journal_res.value;
                }
            }
            if (journal_list.length === 0) {
                journal_res.value = "";
                journal_res.innerHTML = "";
            }
            filter_data_selected.journal_ids = journal_ids;

            // Account Tags
            var account_tag_ids = [];
            var account_tag_text = [];
            var account_tag_res = document.getElementById("acc_tag_res");
            var account_tag_list = $(".account-tag").select2("data");
            for (i = 0; i < account_tag_list.length; i++) {
                if (account_tag_list[i].element[0].selected === true) {
                    account_tag_ids.push(parseInt(account_tag_list[i].id, 10));
                    if (account_tag_text.includes(account_tag_list[i].text) === false) {
                        account_tag_text.push(account_tag_list[i].text);
                    }
                    account_tag_res.value = account_tag_text;
                    account_tag_res.innerHTML = account_tag_res.value;
                }
            }
            if (account_tag_list.length === 0) {
                account_tag_res.value = "";
                account_tag_res.innerHTML = "";
            }
            filter_data_selected.account_tag_ids = account_tag_ids;

            // Analytic Accounts
            var analytic_ids = [];
            var analytic_text = [];
            var analytic_res = document.getElementById("analytic_res");
            var analytic_list = $(".analytics").select2("data");
            for (i = 0; i < analytic_list.length; i++) {
                if (analytic_list[i].element[0].selected === true) {
                    analytic_ids.push(parseInt(analytic_list[i].id, 10));
                    if (analytic_text.includes(analytic_list[i].text) === false) {
                        analytic_text.push(analytic_list[i].text);
                    }
                    analytic_res.value = analytic_text;
                    analytic_res.innerHTML = analytic_res.value;
                }
            }
            if (analytic_list.length === 0) {
                analytic_res.value = "";
                analytic_res.innerHTML = "";
            }
            filter_data_selected.analytic_ids = analytic_ids;

            // Analytic Tags
            var analytic_tag_ids = [];
            var analytic_tag_text = [];
            var analytic_tag_res = document.getElementById("analic_tag_res");
            var analytic_tag_list = $(".analytic-tag").select2("data");
            for (i = 0; i < analytic_tag_list.length; i++) {
                if (analytic_tag_list[i].element[0].selected === true) {
                    analytic_tag_ids.push(parseInt(analytic_tag_list[i].id, 10));
                    if (
                        analytic_tag_text.includes(analytic_tag_list[i].text) === false
                    ) {
                        analytic_tag_text.push(analytic_tag_list[i].text);
                    }
                    analytic_tag_res.value = analytic_tag_text;
                    analytic_tag_res.innerHTML = analytic_tag_res.value;
                }
            }
            if (analytic_tag_list.length === 0) {
                analytic_tag_res.value = "";
                analytic_tag_res.innerHTML = "";
            }
            filter_data_selected.analytic_tag_ids = analytic_tag_ids;

            // Operating Units
            var operating_unit_ids = [];
            var operating_unit_text = [];
            var operating_unit_res = document.getElementById("operating_unit_res");
            var operating_unit_list = $(".operating-units").select2("data");
            for (i = 0; i < operating_unit_list.length; i++) {
                if (operating_unit_list[i].element[0].selected === true) {
                    operating_unit_ids.push(parseInt(operating_unit_list[i].id, 10));
                    if (
                        operating_unit_text.includes(operating_unit_list[i].text) ===
                        false
                    ) {
                        operating_unit_text.push(operating_unit_list[i].text);
                    }
                    operating_unit_res.value = operating_unit_text;
                    operating_unit_res.innerHTML = operating_unit_res.value;
                }
            }
            if (operating_unit_list.length === 0) {
                operating_unit_res.value = "";
                operating_unit_res.innerHTML = "";
            }
            filter_data_selected.operating_unit_ids = operating_unit_ids;

            // Date Range
            var dateString = "";
            if ($("#date_from").val()) {
                dateString = $("#date_from").val();
                filter_data_selected.date_from = dateString;
            }
            if ($("#date_to").val()) {
                dateString = $("#date_to").val();
                filter_data_selected.date_to = dateString;
            }

            // Target Move
            if ($(".target_move").length) {
                var post_res = document.getElementById("post_res");
                filter_data_selected.target_move = $(".target_move")[1].value;
                post_res.value = $(".target_move")[1].value;
                post_res.innerHTML = post_res.value;
                if ($(".target_move")[1].value === "") {
                    post_res.innerHTML = "posted";
                }
            }

            // Debit/Credit
            if ($(".debit_credit").length) {
                var debit_credit_res = document.getElementById("debit_credit_res");
                filter_data_selected.debit_credit = $(".debit_credit")[1].value;
                debit_credit_res.value = $(".debit_credit")[1].value;
                debit_credit_res.innerHTML = debit_credit_res.value;
                if ($(".debit_credit")[1].value === "") {
                    debit_credit_res.innerHTML = "show";
                }
            }
            rpc.query({
                model: "dynamic.balance.sheet.report",
                method: "write",
                args: [self.wizard_id, filter_data_selected],
            }).then(function () {
                self.load_data(self.initial_render);
            });
        },
    });

    core.action_registry.add("dfr_n", ProfitAndLoss);
    return ProfitAndLoss;
});
