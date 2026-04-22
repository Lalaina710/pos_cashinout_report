import io
import base64
from datetime import datetime, time

import pytz

from odoo import fields, models, _
from odoo.exceptions import UserError


class PosCashInOutReportWizard(models.TransientModel):
    _name = 'pos.cashinout.report.wizard'
    _description = 'Wizard Rapport Cash In/Out PdV'

    date_from = fields.Date(
        string='Date début', required=True,
        default=lambda self: fields.Date.today().replace(day=1),
    )
    date_to = fields.Date(
        string='Date fin', required=True,
        default=fields.Date.today,
    )
    pos_config_ids = fields.Many2many(
        'pos.config', string='Point(s) de vente',
        help='Laisser vide pour tous les PdV.',
    )
    movement_type = fields.Selection(
        [('all', 'Tous'), ('in', 'Cash In uniquement'), ('out', 'Cash Out uniquement')],
        string='Type mouvement', default='all', required=True,
    )
    exclude_session_total = fields.Boolean(
        string='Exclure encaissements session',
        default=True,
        help="Exclut les lignes de règlement de session (ref='POS/XXXXX') "
             "pour ne garder que les cash in/out manuels et écarts.",
    )
    exclude_closing_gap = fields.Boolean(
        string='Exclure écarts de clôture',
        default=False,
        help="Exclut les lignes 'Écart d'espèces observé lors du comptage'.",
    )
    show_pos_summary = fields.Boolean(
        string='Synthèse par PdV',
        default=True,
        help='Ajoute un onglet avec totaux IN/OUT et solde par point de vente.',
    )
    report_file = fields.Binary('Fichier', readonly=True)
    report_filename = fields.Char('Nom du fichier', readonly=True)

    def _classify(self, ref, amount):
        """Retourne libellé catégorie et type IN/OUT."""
        ref_low = (ref or '').lower()
        if 'écart' in ref_low or 'ecart' in ref_low:
            return ('Écart clôture', 'in' if amount >= 0 else 'out')
        if ref_low.startswith('pos/') and '-in-' not in ref_low and '-out-' not in ref_low:
            return ('Règlement session', 'in' if amount >= 0 else 'out')
        if '-out-' in ref_low or 'out' in ref_low.split('-'):
            return ('Cash Out', 'out')
        if '-in-' in ref_low or 'in' in ref_low.split('-'):
            return ('Cash In', 'in')
        return ('Autre', 'in' if amount >= 0 else 'out')

    def _get_data(self):
        tz = pytz.timezone(self.env.user.tz or 'Indian/Antananarivo')
        dt_from = tz.localize(datetime.combine(
            self.date_from, time.min,
        )).astimezone(pytz.utc).replace(tzinfo=None)
        dt_to = tz.localize(datetime.combine(
            self.date_to, time.max,
        )).astimezone(pytz.utc).replace(tzinfo=None)

        domain = [
            ('pos_session_id', '!=', False),
            ('move_id.date', '>=', dt_from.date()),
            ('move_id.date', '<=', dt_to.date()),
        ]
        if self.pos_config_ids:
            domain.append(('pos_session_id.config_id', 'in', self.pos_config_ids.ids))

        lines = self.env['account.bank.statement.line'].search(
            domain, order='move_id, id',
        )

        rows = []
        for line in lines:
            amount = line.amount or 0.0
            ref = line.payment_ref or ''
            categ, mvt = self._classify(ref, amount)

            if self.exclude_session_total and categ == 'Règlement session':
                continue
            if self.exclude_closing_gap and categ == 'Écart clôture':
                continue
            if self.movement_type == 'in' and mvt != 'in':
                continue
            if self.movement_type == 'out' and mvt != 'out':
                continue

            session = line.pos_session_id
            rows.append({
                'date': line.move_id.date.strftime('%d/%m/%Y') if line.move_id.date else '',
                'pos_name': session.config_id.name if session.config_id else '',
                'session_name': session.name or '',
                'user_name': session.user_id.name if session.user_id else '',
                'categ': categ,
                'type': 'IN' if mvt == 'in' else 'OUT',
                'ref': ref,
                'amount_in': amount if mvt == 'in' else 0.0,
                'amount_out': abs(amount) if mvt == 'out' else 0.0,
                'amount_signed': amount,
            })
        return rows

    def _get_pos_summary(self, rows):
        summary = {}
        for r in rows:
            pos = r['pos_name'] or 'Non défini'
            if pos not in summary:
                summary[pos] = {
                    'pos_name': pos,
                    'nb': 0,
                    'total_in': 0.0,
                    'total_out': 0.0,
                    'solde': 0.0,
                }
            summary[pos]['nb'] += 1
            summary[pos]['total_in'] += r['amount_in']
            summary[pos]['total_out'] += r['amount_out']
            summary[pos]['solde'] += r['amount_signed']
        result = list(summary.values())
        result.sort(key=lambda x: abs(x['solde']), reverse=True)
        return result

    def action_export_excel(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('La date de début doit être antérieure à la date de fin.'))

        rows = self._get_data()
        pos_summary = self._get_pos_summary(rows) if self.show_pos_summary else None
        content = self._generate_xlsx(rows, pos_summary=pos_summary)
        self.report_file = base64.b64encode(content)
        self.report_filename = 'cash_inout_pdv_%s_%s.xlsx' % (
            self.date_from.strftime('%Y%m%d'),
            self.date_to.strftime('%Y%m%d'),
        )
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=%s&id=%d&field=report_file'
                   '&filename_field=report_filename&download=true' % (
                       self._name, self.id),
            'target': 'new',
        }

    def _generate_xlsx(self, rows, pos_summary=None):
        import xlsxwriter

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        fmt_title = wb.add_format({'bold': True, 'font_size': 14, 'align': 'center'})
        fmt_header = wb.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_header_green = wb.add_format({
            'bold': True, 'bg_color': '#00B050', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_header_red = wb.add_format({
            'bold': True, 'bg_color': '#C00000', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_text = wb.add_format({'border': 1, 'font_size': 10})
        fmt_num = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00',
        })
        fmt_num_neg = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00', 'font_color': 'red',
        })
        fmt_total_lbl = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 11,
        })
        fmt_total_num = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 11, 'num_format': '#,##0.00',
        })

        ws = wb.add_worksheet('Détail Cash In-Out')

        headers = [
            ('DATE', fmt_header, 12),
            ('POINT DE VENTE', fmt_header, 22),
            ('SESSION', fmt_header, 14),
            ('UTILISATEUR', fmt_header, 22),
            ('CATÉGORIE', fmt_header, 20),
            ('TYPE', fmt_header, 8),
            ('MOTIF', fmt_header, 50),
            ('CASH IN', fmt_header_green, 14),
            ('CASH OUT', fmt_header_red, 14),
        ]
        for i, (_lbl, _fmt, w) in enumerate(headers):
            ws.set_column(i, i, w)

        last_col = len(headers) - 1
        ws.merge_range(0, 0, 0, last_col, 'Mouvements Caisse Point de Vente', fmt_title)
        ws.write(1, 0, self.env.company.name, fmt_text)
        ws.write(1, last_col - 1, 'Période', fmt_text)
        ws.write(1, last_col, '%s au %s' % (
            self.date_from.strftime('%d/%m/%Y'),
            self.date_to.strftime('%d/%m/%Y'),
        ), fmt_text)

        row = 3
        for col, (label, fmt, _w) in enumerate(headers):
            ws.write(row, col, label, fmt)
        row += 1
        ws.freeze_panes(row, 0)
        first_data = row

        total_in = 0.0
        total_out = 0.0
        for r in rows:
            ws.write(row, 0, r['date'], fmt_text)
            ws.write(row, 1, r['pos_name'], fmt_text)
            ws.write(row, 2, r['session_name'], fmt_text)
            ws.write(row, 3, r['user_name'], fmt_text)
            ws.write(row, 4, r['categ'], fmt_text)
            ws.write(row, 5, r['type'], fmt_text)
            ws.write(row, 6, r['ref'], fmt_text)
            ws.write(row, 7, r['amount_in'], fmt_num)
            ws.write(row, 8, r['amount_out'], fmt_num)
            total_in += r['amount_in']
            total_out += r['amount_out']
            row += 1

        if row > first_data:
            ws.autofilter(first_data - 1, 0, row - 1, last_col)

        solde = total_in - total_out
        ws.merge_range(row, 0, row, 6, 'TOTAL', fmt_total_lbl)
        ws.write(row, 7, total_in, fmt_total_num)
        ws.write(row, 8, total_out, fmt_total_num)
        row += 1
        ws.merge_range(row, 0, row, 7, 'SOLDE NET (IN - OUT)', fmt_total_lbl)
        ws.write(row, 8, solde, fmt_total_num if solde >= 0 else fmt_num_neg)

        if pos_summary:
            ws2 = wb.add_worksheet('Synthèse par PdV')
            ws2.merge_range(0, 0, 0, 4, 'Synthèse Cash In/Out par PdV', fmt_title)
            ws2.write(1, 0, self.env.company.name, fmt_text)
            ws2.write(1, 3, 'Période', fmt_text)
            ws2.write(1, 4, '%s au %s' % (
                self.date_from.strftime('%d/%m/%Y'),
                self.date_to.strftime('%d/%m/%Y'),
            ), fmt_text)

            s_headers = [
                ('POINT DE VENTE', fmt_header, 28),
                ('NB LIGNES', fmt_header, 12),
                ('TOTAL IN', fmt_header_green, 16),
                ('TOTAL OUT', fmt_header_red, 16),
                ('SOLDE NET', fmt_header, 16),
            ]
            for i, (_lbl, _fmt, w) in enumerate(s_headers):
                ws2.set_column(i, i, w)
            s_row = 3
            for col, (label, fmt, _w) in enumerate(s_headers):
                ws2.write(s_row, col, label, fmt)
            s_row += 1
            ws2.freeze_panes(s_row, 0)

            s_total_nb = 0
            s_total_in = 0.0
            s_total_out = 0.0
            for r in pos_summary:
                ws2.write(s_row, 0, r['pos_name'], fmt_text)
                ws2.write(s_row, 1, r['nb'], fmt_num)
                ws2.write(s_row, 2, r['total_in'], fmt_num)
                ws2.write(s_row, 3, r['total_out'], fmt_num)
                ws2.write(s_row, 4, r['solde'], fmt_num if r['solde'] >= 0 else fmt_num_neg)
                s_total_nb += r['nb']
                s_total_in += r['total_in']
                s_total_out += r['total_out']
                s_row += 1
            ws2.autofilter(3, 0, s_row - 1, 4)
            s_solde = s_total_in - s_total_out
            ws2.write(s_row, 0, 'TOTAL', fmt_total_lbl)
            ws2.write(s_row, 1, s_total_nb, fmt_total_num)
            ws2.write(s_row, 2, s_total_in, fmt_total_num)
            ws2.write(s_row, 3, s_total_out, fmt_total_num)
            ws2.write(s_row, 4, s_solde, fmt_total_num if s_solde >= 0 else fmt_num_neg)

        wb.close()
        return output.getvalue()
