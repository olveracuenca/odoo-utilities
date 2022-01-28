# Copyright 2019, Jarsa Sistemas, S.A. de C.V.
# License LGPL-3.0 or later (http://www.gnu.org/licenses/lgpl.html).
# pylint: disable=W7936

import logging
import re
import string

from googleapiclient import discovery
from googleapiclient.errors import HttpError
import google.oauth2.credentials
from odoo import _, api, fields, models
from odoo.exceptions import UserError

_logger = logging.getLogger(__name__)
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)


class GoogleDriveFile(models.Model):
    _name = 'google.spreadsheet.file'
    _description = 'Spreadsheet file to be imported'
    _order = 'sequence'

    name = fields.Char()
    file_type = fields.Selection([
        ('import', 'Import Spreadsheet'),
        ('create', 'Create Spreadsheet')],
        required=True, default='import')
    active = fields.Boolean(default=True)
    sequence = fields.Integer(default=10)
    url = fields.Char(
        string="URL",
        help='Technical field used to save the URL of the file and allow to '
        'open the file from Odoo.',
        readonly=True)
    id_file = fields.Char(
        string='Spreadsheet ID',
        help='ID of the spreadsheet. If the URL of the spreadsheet is '
        'https://docs.google.com/spreadsheets/d/1SBcVQu/edit#gid=123 the '
        'ID will be 1SBcVQu')
    sheet_ids = fields.One2many(
        'google.spreadsheet.file.sheet', 'file_id',
        string='Sheets', readonly=True)
    model_ids = fields.One2many(
        'google.spreadsheet.file.model', 'file_id',
        string='Models')

    @api.model
    def _get_service(self):
        access_token = self.env['google.drive.config'].get_access_token()
        credentials = google.oauth2.credentials.Credentials(access_token)
        return discovery.build(
            'sheets', 'v4', credentials=credentials)

    @api.model
    def _extract_id_from_url(self, url):
        res = re.compile(r'/spreadsheets/d/([a-zA-Z0-9-_]+)').search(url)
        if res:
            return res.group(1)
        res = re.compile(r'key=([^&#]+)').search(url)
        if res:
            return res.group(1)
        raise UserError(_('The URL provided is invalid.'))

    def get_file_info(self):
        for rec in self:
            service = rec._get_service()
            try:
                sheet_metadata = service.spreadsheets().get(
                    spreadsheetId=self.id_file).execute()
            except HttpError:
                raise UserError(_(
                    'The Spreadsheet ID is not correct, please verify it.'))
            vals = rec._get_file_info(sheet_metadata)
            rec.write(vals)

    @api.onchange('id_file')
    def _onchange_id_file(self):
        if self.id_file:
            self.id_file = self._extract_id_from_url(self.id_file)

    def _get_file_info(self, sheet_metadata):
        self.ensure_one()
        vals = {
            'name': sheet_metadata.get('properties', {}).get('title'),
            'url': sheet_metadata.get('spreadsheetUrl', ''),
            'id_file': sheet_metadata.get('spreadsheetId', ''),
            'sheet_ids': [],
        }
        sheet_ids = []
        for sheet in sheet_metadata.get('sheets', []):
            id_sheet = sheet.get("properties", {}).get("sheetId", 0)
            sheet_ids.append(id_sheet)
            name = sheet.get("properties", {}).get('title', "Sheet1")
            if id_sheet in self.sheet_ids.mapped('id_sheet'):
                sheet_id = self.sheet_ids.filtered(
                    lambda s: s.id_sheet == id_sheet)
                vals['sheet_ids'].append((1, sheet_id.id, {
                    'name': name,
                }))
                continue
            vals['sheet_ids'].append((0, 0, {
                'name': name,
                'id_sheet': id_sheet,
                'sequence': sheet.get("properties", {}).get("index", 0),
            }))
        self.sheet_ids.filtered(
            lambda s: s.id_sheet not in sheet_ids).unlink()
        return vals

    @api.model
    def _get_range(self, sheet_name, values):
        alphabet = list(string.ascii_uppercase)
        alphabet.extend([i + b for i in alphabet for b in alphabet])
        return '%s!A1:%s1' % (sheet_name, alphabet[len(values) - 1])

    @api.model
    def _get_invalid_fields(self):
        return [
            'activity_ids',
            'activity_summary',
            'activity_type_id',
            'activity_user_id',
            'message_follower_ids',
            'message_ids',
            'message_main_attachment_id',
            'signup_expiration',
            'signup_token',
            'signup_type',
        ]

    def create_update_file(self):
        for rec in self:
            spreadsheet = {
                'properties': {
                    'title': rec.name,
                    'locale': self._context.get('lang', 'en_US'),
                    'autoRecalc': 'ON_CHANGE',
                    'timeZone': (
                        self._context.get(
                            'tz') if self._context.get(
                            'tz') else 'America/New_York'),
                },
            }
            sheets_data = {
                'valueInputOption': 'RAW',
                'data': [],
            }
            requests = []
            sheet_id = 0
            for model in rec.model_ids.mapped('model_id'):
                spreadsheet.setdefault('sheets', []).append({
                    'properties': {
                        'sheetId': sheet_id,
                        'title': '%s(%s)' % (model.name, model.model),
                        'hidden': False,
                    },
                })
                sheets = spreadsheet['sheets'][sheet_id]
                sheets.setdefault('data', []).append({
                    'rowData': [{
                        'values': []
                    }],
                })
                values = sheets['data'][0]['rowData'][0]['values']
                names = [[]]
                invalid_fields = self._get_invalid_fields()
                for field in model.field_id.filtered(
                        lambda f: f.readonly is False and
                        f.name not in invalid_fields):
                    names[0].append(field.name)
                    values.append({
                        'note': '%s\n%s' % (
                            field.field_description, field.help or ''),
                    })
                sheets_data['data'].append({
                    'range': self._get_range(
                        '%s(%s)' % (model.name, model.model), names[0]),
                    'majorDimension': 'ROWS',
                    'values': names,
                })
                requests.append({
                    'updateSheetProperties': {
                        'properties': {
                            'sheetId': sheet_id,
                            'gridProperties': {
                                'frozenRowCount': 1,
                            }
                        },
                        'fields': 'gridProperties.frozenRowCount',
                    }
                })
                sheet_id += 1
            service = rec._get_service()
            sheet_metadata = service.spreadsheets().create(
                body=spreadsheet).execute()
            service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_metadata.get('spreadsheetId'),
                body=sheets_data).execute()
            service.spreadsheets().batchUpdate(
                spreadsheetId=sheet_metadata.get('spreadsheetId'),
                body={'requests': requests}).execute()
            vals = rec._get_file_info(sheet_metadata)
            rec.write(vals)

    def open_file(self):
        self.ensure_one()
        return {
            'type': 'ir.actions.act_url',
            'url': self.url,
        }


class GoogleDriveFileSheet(models.Model):
    _name = 'google.spreadsheet.file.sheet'
    _description = 'Sheets of Spreadsheet File'
    _order = 'sequence'

    file_id = fields.Many2one(
        'google.spreadsheet.file', string='File', readonly=True,
        ondelete='cascade')
    name = fields.Char(readonly=True)
    id_sheet = fields.Integer(string='Sheet ID', readonly=True)
    sequence = fields.Integer(string='Index', readonly=True)


class GoogleDriveFileModel(models.Model):
    _name = 'google.spreadsheet.file.model'
    _description = 'Models to create a template in Google Spreadsheets'
    _order = 'sequence'

    file_id = fields.Many2one(
        'google.spreadsheet.file', string='File', readonly=True,
        ondelete='cascade')
    sequence = fields.Integer(default=10)
    model_id = fields.Many2one(
        comodel_name='ir.model', string='Relation model', required=True, ondelete='cascade')
    model = fields.Char(related='model_id.model', readonly=True)
