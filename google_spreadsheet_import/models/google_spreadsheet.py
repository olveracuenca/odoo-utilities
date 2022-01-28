# Copyright 2019, Jarsa Sistemas, S.A. de C.V.
# License LGPL-3.0 or later (http://www.gnu.org/licenses/lgpl.html).
# pylint: disable=redefined-builtin
# pylint: disable=deprecated-module
# pylint: disable=too-complex
# pylint: disable=missing-manifest-dependency

import logging
from csv import DictReader
from io import StringIO
from pytz import timezone
import datetime

import requests
from odoo import _, api, fields, models
from odoo.exceptions import ValidationError, Warning
from odoo.tools.safe_eval import safe_eval, test_python_expr, wrap_module
from odoo import tools

_logger = logging.getLogger(__name__)

tools.safe_eval.ipdb = wrap_module(__import__('ipdb'), ['set_trace'])
base64 = wrap_module(__import__('base64'), ['b64encode', 'b64decode'])
json = wrap_module(__import__('json'), ['loads', 'dumps'])
math = wrap_module(__import__('math'), ['isclose'])
re = wrap_module(__import__('re'), [
    'match', 'findall', 'compile', 'search', 'finditer'])
requests2 = wrap_module(__import__('requests'), ['get'])


class GoogleDriveSheet(models.Model):
    _name = 'google.spreadsheet'
    _description = 'Sync google drive sheet'
    _order = 'sequence'

    def _get_default_python_code(self):
        code = self.env['ir.actions.server'].DEFAULT_PYTHON_CODE
        code = code[:-3]
        code += (
            '# To create a External ID use this code:'
            'create_external_id(record, xml_id)\n'
            '# To log a new error in this record, use this code: log_error('
            'message, field=False, record=False,type=\'error\')')
        return code

    name = fields.Char(required=True)
    active = fields.Boolean(default=True)
    sequence = fields.Integer(default=10)
    model_id = fields.Many2one(
        comodel_name='ir.model', string='Relation model', required=True,
        ondelete='cascade')
    model = fields.Char(related='model_id.model', readonly=True)
    file_id = fields.Many2one('google.spreadsheet.file', required=True)
    sheet_id = fields.Many2one(
        'google.spreadsheet.file.sheet',
        help='Sheet to be used, if is not provided the first '
        'sheet of the file will be used')
    sheet_range = fields.Char(
        help='Used to get only the information in the range provided.\n'
        'Example: A1:G')
    query = fields.Char(
        help="Used to get data from Google Spreadsheets using a Query:\n"
        "Example: select A, B, (D+E) where C < 100 and X = 'yes'\n"
        "For more information visit:\n"
        "https://developers.google.com/chart/interactive/docs/querylanguage")
    error_ids = fields.One2many(
        'google.spreadsheet.error', 'sheet_id',
        string='List of errors', readonly=True)
    log_ids = fields.One2many(
        'google.spreadsheet.log', 'sheet_id',
        string='List of logs', readonly=True)
    batch_size = fields.Integer(
        help='Used to define the size of the batch of records that will be '
        'imported at the same time, this helps to prevent CPU timeouts.',
        default=50, required=True)
    date_format = fields.Char(default='%d-%m-%Y')
    datetime_format = fields.Char(default='%d-%m-%Y %H:%M:%S')
    separator = fields.Char(default=',')
    float_decimal_separator = fields.Char(default='.')
    encoding = fields.Char(default='utf-8')
    quoting = fields.Char(default='"')
    float_thousand_separator = fields.Char(default=',')
    context = fields.Char(
        help='Technical field used to pass context to importation process.',
        default={})
    code = fields.Text(
        string='Python Code',
        default=_get_default_python_code,
        help='Write Python code that the action will execute. Some variables '
             'are available for use; help about python expression is given in '
             'the help tab.')
    import_type = fields.Selection([
        ('native', 'Native'),
        ('code', 'Python Code')], default='native',
        help='Native: Use the spreadsheet with odoo native importation.\n'
             'Code: Write python code to manipulate the data from the '
             'spreadsheet.')
    group_ids = fields.Many2many(
        'res.groups', string='Groups',
        help='Define what groups will have access to import this sheet if this'
             ' field is empty only the administrator will have the access to '
             'this sheet.')
    user_ids = fields.Many2many(
        'res.users', string='Users',
        help='Define what users that will have access to import this sheet if '
             'this field is empty only the administrator will have the access '
             'to this sheet.')
    fix_header = fields.Boolean(
        help='In some cases the header of the document is not taken correctly.'
             ' Check this field to fix the header.')
    header_value = fields.Integer(default=0)
    data = fields.Text(readonly=True)
    store_data = fields.Text(readonly=True)
    background_import = fields.Boolean(readonly=True)

    def _get_content(self, id_file=""):
        self.ensure_one()
        access_token = self.env['google.drive.config'].get_access_token()
        params = {
            'access_token': access_token,
            'tqx': 'out:csv',
        }
        if self.fix_header:
            params['headers'] = self.header_value
        if self.sheet_id:
            params['sheet'] = self.sheet_id.name
        if self.sheet_range:
            params['range'] = self.sheet_range
        if self.query:
            params['tq'] = self.query
        url = 'https://docs.google.com/spreadsheets/d/%s/gviz/tq' % id_file
        return requests.get(url, params=params)

    @api.model
    def _split_list(self, records, number):
        """ This function split a list and retuns a list of lists.
        Example:
        list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        number = 5
        returns [[1, 2, 3, 4, 5], [6, 7, 8, 9, 10]]
        """
        def chunks(records, number):
            for index in range(0, len(records), number):
                yield records[index:index+number]
        return list(chunks(records, number))

    def action_open_native_import(self):
        self.ensure_one()
        context = {}
        if self.context:
            try:
                context = safe_eval(self.context)
            except ValueError:
                raise ValidationError(_(
                    'The context must be formatted as python dictionary.'))
        return {
            'type': 'ir.actions.client',
            'tag': 'import',
            'params': {
                'model': self.model,
                'context': context,
            }
        }

    def clean_log(self):
        for rec in self:
            rec.error_ids.unlink()

    @api.constrains('code')
    def _check_python_code(self):
        for rec in self.filtered('code'):
            msg = test_python_expr(expr=rec.code.strip(), mode="exec")
            if msg:
                raise ValidationError(msg)

    def upload(self):
        for rec in self:
            if rec.background_import:
                data = rec.data
            else:
                response = self._get_content(rec.file_id.id_file)
                if response.status_code != 200:
                    raise ValidationError(_(
                        'There was an error, please contact the Administrator')
                    )
                data = response.text
            data = rec._process_data(data)
            time_start = datetime.datetime.now()
            res = getattr(self, '_process_%s' % rec.import_type)(data)
            name = 'Records created: %s' % len(res.get('ids'))
            if res.get('errors'):
                name += '\nErrors: %s' % len(res.get('errors'))
            time_end = datetime.datetime.now()
            rec.write({
                'store_data': res.get('store_data'),
                'log_ids': [(0, 0, {
                    'name': name,
                    'duration': str(time_end - time_start),
                    'ids_related': (
                        ','.join(str(id) for id in res.get('ids')) or ''),
                })]
            })
            return res.get('action')

    def _get_eval_context(self, records):
        """ Prepare the context used when evaluating python code, like the
        python formulas or code server actions.

        :param records: dictionary with google spreadsheet data
        :type DictReader: records to be processed
        :returns: dict -- evaluation context given to (safe_)safe_eval """
        self.ensure_one()

        def log(message, level="info"):
            with self.pool.cursor() as cr:
                cr.execute("""
                    INSERT INTO ir_logging(
                    create_date, create_uid, type, dbname, name, level,
                    message, path, line, func)
                    VALUES (
                    NOW() at time zone 'UTC',
                    %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    self.env.uid, 'server', self._cr.dbname, __name__, level,
                    message, "google spreadsheet", self.id, self.name))

        def log_error(message, field=False, record=False, type='error'):
            self.write({
                'error_ids': [(0, 0, {
                    'field': field,
                    'message': message,
                    'record': record,
                    'type': type,
                })],
            })

        def create_external_id(record, xml_id):
            self.env['ir.model.data'].create({
                'module': '__import__',
                'model': record._name,
                'name': xml_id,
                'res_id': record.id,
            })
        return {
            'records': records,
            'datetime': tools.safe_eval.datetime,
            'dateutil': tools.safe_eval.dateutil,
            'time': tools.safe_eval.time,
            'timezone': timezone,
            'uid': self.env.uid,
            'user': self.env.user,
            'env': self.env,
            'model': self.env[self.model],
            'Warning': Warning,
            'log': log,
            'ipdb': tools.safe_eval.ipdb,
            'math': math,
            'create_external_id': create_external_id,
            'log_error': log_error,
            'base64': base64,
            'json': json,
            're': re,
            'requests': requests2,
            'store_data': (
                self.store_data and json.loads(self.store_data) or {}),
        }

    def activate_background_import(self):
        for rec in self:
            response = self._get_content(rec.file_id.id_file)
            if response.status_code != 200:
                raise ValidationError(_(
                    'There was an error, please contact the Administrator')
                )
            data = rec._process_data(response.text)
            data['records'] = self._split_list(
                data['records'], self.batch_size)
            rec.write({
                'data': json.dumps(data, indent=2),
                'background_import': True,
            })

    def deactivate_background_import(self):
        self.write({
            'data': False,
            'background_import': False
        })

    def _process_background_import(self):
        sheets = self.search([('background_import', '=', True)])
        for sheet in sheets:
            sheet.upload()
            data = json.loads(sheet.data)
            if not data['records']:
                sheet.write({
                    'data': False,
                    'background_import': False,
                })

    def _process_data(self, data):
        self.ensure_one()
        if self.background_import:
            data = json.loads(self.data)
            if data['records']:
                records = data['records'].pop(0)
                self.data = json.dumps(data, indent=2)
                data['records'] = records
            else:
                self.write({
                    'data': False,
                    'background_import': False,
                })
            return data
        if self.import_type == 'native':
            records = data.split('\n')
            header_str = records.pop(0)
            header = header_str.replace('"', '').split(',')
            columns = header
            model_fields = self.model_id.mapped('field_id.name')
            for idx, field in enumerate(header):
                clean_field = field.split('/')[0]
                if clean_field not in model_fields:
                    header[idx] = False
            return {
                'records': records,
                'columns': columns,
                'header': header,
            }
        if self.import_type == 'code':
            return {
                'records': [dict(rec) for rec in DictReader(StringIO(data))],
            }

    def _process_code(self, data):
        for rec in self:
            records = data['records']
            _logger.info(
                'model %s importing %d rows...', self.model, len(records))
            eval_context = self._get_eval_context(records)
            safe_eval(
                rec.code.strip(), eval_context, mode="exec", nocopy=True)
            return {
                'errors': eval_context.get('errors', []),
                'ids': eval_context.get('ids', ''),
                'action': eval_context.get('action', {}),
                'store_data': json.dumps(
                    eval_context.get('store_data', {}), indent=2)
            }

    def _process_native(self, data):
        for rec in self:
            rec.error_ids.unlink()
            records = data['records']
            records = self._split_list(records, self.batch_size)
            header = data['header']
            columns = data['columns']
            base_import = self.env['base_import.import']
            thousand = rec.float_thousand_separator
            decimal = rec.float_decimal_separator
            default_options = {
                'separator': rec.separator if rec.separator else ',',
                'headers': True,
                'fields': [],
                'float_decimal_separator': decimal if decimal else '.',
                'encoding': rec.encoding if rec.encoding else 'utf-8',
                'quoting': rec.quoting if rec.quoting else '"',
                'float_thousand_separator': thousand if thousand else ',',
                'date_format': rec.date_format if rec.date_format else '',
                'advanced': True,
                'keep_matches': False,
                'datetime_format': (
                    rec.datetime_format if rec.datetime_format else ''),
            }
            context = {}
            if rec.context:
                try:
                    context = safe_eval(rec.context)
                except ValueError:
                    raise ValidationError(_(
                        'The context must be formatted as python dictionary.'))
            errors = []
            ids = []
            count = 0
            for record in records:
                new_import = base_import.create({
                    'file': '\n'.join(record).encode('utf-8'),
                    'res_model': rec.model,
                    'file_name': 'google_export.csv',
                    'file_type': 'text/csv; charset=utf-8',
                })
                do = new_import.with_context(**context).execute_import(
                    header, columns, default_options)
                if do.get('messages'):
                    for error in do.get('messages'):
                        error['sheet_id'] = rec.id
                        # We need to add 2 because Odoo return row number
                        # starting by 0 and we also need to add the header
                        # line to get the correct row number.
                        row_num = count * rec.batch_size + 2
                        if error.get('rows'):
                            error['row_from'] = error['rows']['from'] + row_num
                            error['row_to'] = error['rows']['to'] + row_num
                            error.pop('rows')
                        if error.get('moreinfo'):
                            error.pop('moreinfo')
                        if error.get('field_path'):
                            error.pop('field_path')
                        if error.get('field_type'):
                            error.pop('field_type')
                        if error.get('record', False):
                            error['record'] = error['record'] + row_num
                        errors.append(error)
                if do.get('ids'):
                    ids.extend(do.get('ids'))
                count += 1
            if errors:
                rec.error_ids.create(errors)
            # TODO Create a wizard to check if have combined stuff
            # errors and ids
            return {
                'errors': errors,
                'ids': ids,
                'action': {
                    'name': rec.model_id.display_name,
                    'type': 'ir.actions.act_window',
                    'view_mode': 'tree,form',
                    'res_model': rec.model,
                    'domain': [('id', 'in', ids)],
                }
            }

    def open_file(self):
        self.ensure_one()
        url = '%s#gid=%s' % (self.file_id.url, self.sheet_id.id_sheet)
        return {
            'type': 'ir.actions.act_url',
            'url': url,
        }


class GoogleDriveSheetError(models.Model):
    _name = 'google.spreadsheet.error'
    _description = 'Errors of Google Drive Sheets'

    sheet_id = fields.Many2one(
        'google.spreadsheet', readonly=True, ondelete='cascade')
    field = fields.Char(readonly=True)
    field_name = fields.Char(readonly=True)
    message = fields.Char(readonly=True)
    record = fields.Char(readonly=True)
    row_from = fields.Char(readonly=True)
    row_to = fields.Char(readonly=True)
    type = fields.Selection([('error', 'Error'), ('warning', 'Warning')])
    value = fields.Char(readonly=True)


class GoogleDriveSheetLog(models.Model):
    _name = 'google.spreadsheet.log'
    _description = 'Log of Google Drive Sheets'
    _order = 'create_date desc'

    sheet_id = fields.Many2one(
        'google.spreadsheet', readonly=True, ondelete='cascade')
    name = fields.Char(readonly=True)
    duration = fields.Char(readonly=True)
    ids_related = fields.Char(
        readonly=True,
        help='Technical field used to save the ids of the records that was '
        'updated/created by this sheet')

    def action_open_related_records(self):
        self.ensure_one()
        ids = [int(id) for id in self.ids_related.split(',')]
        return {
            'name': self.sheet_id.model_id.display_name,
            'type': 'ir.actions.act_window',
            'view_mode': 'tree,form',
            'res_model': self.sheet_id.model,
            'domain': [('id', 'in', ids)],
        }
