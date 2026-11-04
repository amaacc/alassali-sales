"""
odoo_client.py — Odoo 17 Enterprise SaaS Integration via XML-RPC (Strictly Read-Only)
Features:
- Safe Read-Only operations (search_read only, never write or execute changes).
- Smart Morning Radar: Evaluates open and closed sessions, distinguishing between main daily
  sessions and permitted secondary monthly sessions, highlighting late showrooms.
- Direct Sales Data Pull: Extracts financial breakdowns directly from Odoo pos.session & pos.order.
"""

import os
import xmlrpc.client
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Any, Optional, Set, Tuple

# Saudi Arabia is UTC+3
KSA_OFFSET = timezone(timedelta(hours=3))

# Resilient Socket Timeout Transports for XML-RPC (Prevents hanging/freezing)
class SafeTimeoutTransport(xmlrpc.client.SafeTransport):
    def __init__(self, timeout: float = 25.0, **kwargs):
        super().__init__(**kwargs)
        self.timeout = timeout

    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = self.timeout
        return conn

class RegularTimeoutTransport(xmlrpc.client.Transport):
    def __init__(self, timeout: float = 25.0, **kwargs):
        super().__init__(**kwargs)
        self.timeout = timeout

    def make_connection(self, host):
        conn = super().make_connection(host)
        conn.timeout = self.timeout
        return conn

def get_xmlrpc_transport(url: str, timeout: float = 25.0):
    if str(url).lower().startswith("https://"):
        return SafeTimeoutTransport(timeout=timeout)
    return RegularTimeoutTransport(timeout=timeout)

# Optional fallback to root security.py credentials if available
try:
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from security import get_odoo_credentials
except Exception:
    get_odoo_credentials = None


class OdooClient:
    def __init__(
        self,
        url: Optional[str] = None,
        db: Optional[str] = None,
        username: Optional[str] = None,
        password: Optional[str] = None,
        timeout: float = 25.0
    ):
        sec_url, sec_db, sec_user, sec_pwd = "", "", "", ""
        if get_odoo_credentials:
            try:
                sec_url, sec_db, sec_user, sec_pwd = get_odoo_credentials()
            except Exception:
                pass

        self.url = (url or os.environ.get("ODOO_URL", sec_url)).rstrip("/")
        self.db = db or os.environ.get("ODOO_DB", sec_db)
        self.username = username or os.environ.get("ODOO_USERNAME", sec_user)
        self.password = password or os.environ.get("ODOO_PASSWORD", sec_pwd)
        self.timeout = float(timeout or os.environ.get("ODOO_TIMEOUT", 25.0))
        self.uid = None
        self._common = None
        self._models = None

    def connect(self) -> bool:
        """Authenticates with Odoo using XML-RPC with resilient socket timeouts."""
        if not (self.url and self.db and self.username and self.password):
            return False
        try:
            tr_common = get_xmlrpc_transport(self.url, timeout=self.timeout)
            self._common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common", transport=tr_common, allow_none=True)
            self.uid = self._common.authenticate(self.db, self.username, self.password, {})
            if self.uid:
                tr_models = get_xmlrpc_transport(self.url, timeout=self.timeout)
                self._models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object", transport=tr_models, allow_none=True)
                return True
            return False
        except Exception as e:
            print(f"[OdooClient] Connection failed: {e}")
            return False

    def is_connected(self) -> bool:
        return self.uid is not None and self._models is not None

    def search_read(self, model: str, domain: List, fields: List[str], limit: int = 0) -> List[Dict[str, Any]]:
        """Safe read-only wrapper around Odoo search_read."""
        if not self.is_connected() and not self.connect():
            raise ConnectionError("Could not connect to Odoo server. Check credentials.")
        try:
            return self._models.execute_kw(
                self.db, self.uid, self.password,
                model, 'search_read',
                [domain],
                {'fields': fields, 'limit': limit} if limit > 0 else {'fields': fields}
            )
        except Exception as e:
            print(f"[OdooClient] search_read on {model} failed: {e}")
            raise e

    # =========================================================================
    # TAB 1: SMART MORNING RADAR LOGIC
    # =========================================================================
    def get_morning_radar_status(self, old_closings: Set[str]) -> Dict[str, Any]:
        """
        Executes Smart Morning Radar:
        1. Open sessions: state != 'closed' opened before yesterday 17:00 KSA.
        2. Closed unprinted sessions: state == 'closed' and name not in old_closings.
        3. Dynamic main vs secondary classification per showroom.
        """
        now_ksa = datetime.now(KSA_OFFSET)
        yesterday_ksa = now_ksa - timedelta(days=1)
        # Yesterday at 17:00 KSA = yesterday 14:00 UTC
        yesterday_5pm_ksa = yesterday_ksa.replace(hour=17, minute=0, second=0, microsecond=0)
        yesterday_5pm_utc = yesterday_5pm_ksa.astimezone(timezone.utc)
        cutoff_utc_str = yesterday_5pm_utc.strftime("%Y-%m-%d %H:%M:%S")

        # 1. Fetch relevant sessions (opened in last 4 days - optimized window)
        four_days_ago_utc = (now_ksa - timedelta(days=4)).astimezone(timezone.utc).strftime("%Y-%m-%d 00:00:00")
        domain = [('start_at', '>=', four_days_ago_utc)]
        fields = [
            'id', 'name', 'state', 'config_id', 'start_at', 'stop_at',
            'cash_register_total_entry_encoding', 'total_payments_amount', 'order_count'
        ]

        raw_sessions = self.search_read('pos.session', domain, fields)

        # Batch query live order totals for overdue open sessions (Odoo total_payments_amount can be 0.0 for open sessions)
        overdue_open_ids = [
            sess['id'] for sess in raw_sessions
            if sess.get('state') != 'closed' and sess.get('start_at') and sess['start_at'] < cutoff_utc_str
        ]
        orders_by_session = {}
        if overdue_open_ids:
            try:
                orders = self.search_read(
                    'pos.order',
                    [('session_id', 'in', overdue_open_ids), ('state', 'in', ['paid', 'done', 'invoiced'])],
                    ['session_id', 'amount_total']
                )
                for o in orders:
                    sid = o['session_id'][0] if isinstance(o['session_id'], (list, tuple)) else o['session_id']
                    orders_by_session[sid] = orders_by_session.get(sid, 0.0) + float(o.get('amount_total', 0.0) or 0.0)
            except Exception as e:
                print(f"[OdooClient] Failed to fetch live orders for open sessions: {e}")

        # Group by POS Config / Showroom
        by_showroom: Dict[str, Dict[str, Any]] = {}

        for sess in raw_sessions:
            s_name = sess.get('name')
            state = sess.get('state')
            config_tuple = sess.get('config_id')
            pos_name = config_tuple[1] if isinstance(config_tuple, (list, tuple)) else "Unknown POS"
            start_at_str = sess.get('start_at', '')
            stop_at_str = sess.get('stop_at', '')
            
            # Accurate sales amount: use orders_by_session if available, else total_payments_amount
            sales_amount = orders_by_session.get(sess['id'], float(sess.get('total_payments_amount', 0.0) or 0.0))
            order_count = int(sess.get('order_count', 0) or 0)

            # Determine if open session is overdue (opened before yesterday 17:00 KSA)
            is_overdue = False
            if state != 'closed' and start_at_str and start_at_str < cutoff_utc_str:
                is_overdue = True

            is_printed = s_name in old_closings

            if pos_name not in by_showroom:
                by_showroom[pos_name] = {
                    "pos_name": pos_name,
                    "open_sessions": [],
                    "closed_unprinted": [],
                    "status": "GREEN", # GREEN, BLUE, RED
                    "status_text": "جاهز تماماً",
                    "main_session": None,
                    "secondary_sessions": []
                }

            sess_info = {
                "id": sess['id'],
                "name": s_name,
                "state": state,
                "start_at": start_at_str,
                "stop_at": stop_at_str,
                "amount": sales_amount,
                "order_count": order_count,
                "is_overdue": is_overdue,
                "is_printed": is_printed
            }

            if state == 'closed':
                if not is_printed:
                    by_showroom[pos_name]["closed_unprinted"].append(sess_info)
            else:
                by_showroom[pos_name]["open_sessions"].append(sess_info)

        # Evaluate Showrooms according to the Master Prompt algorithm
        # Check if today is near end of month (last 2 days of month)
        next_month = now_ksa.replace(day=28) + timedelta(days=4)
        last_day_of_month = next_month - timedelta(days=next_month.day)
        days_left_in_month = (last_day_of_month.date() - now_ksa.date()).days
        is_month_end_alert = days_left_in_month <= 2

        ready_showrooms = []
        late_showrooms = []
        secondary_open_showrooms = []

        for pos_name, data in by_showroom.items():
            closed_unprinted = data["closed_unprinted"]
            open_sessions = data["open_sessions"]
            overdue_open = [s for s in open_sessions if s["is_overdue"]]

            # Sort closed unprinted by stop_at descending (latest closed session first)
            closed_unprinted.sort(key=lambda s: s.get("stop_at") or "", reverse=True)

            # Check if any overdue open session is a serious/main operational session
            has_serious_open = False
            serious_sessions = []
            for s in overdue_open:
                # A session is a MAIN session if:
                # 1. order_count > 5 OR amount >= 1000.0
                # 2. OR if there's an unprinted closed session, but this open session started AFTER that closed session ended (subsequent full-day shift)
                is_subsequent = False
                if closed_unprinted and closed_unprinted[0].get("stop_at"):
                    if s.get("start_at") and s["start_at"] >= closed_unprinted[0]["stop_at"] and s.get("order_count", 0) > 0:
                        is_subsequent = True

                is_major = (s["amount"] >= 1000.0) or (s.get("order_count", 0) > 5) or is_subsequent

                if is_major:
                    has_serious_open = True
                    serious_sessions.append(s)
                else:
                    data["secondary_sessions"].append(s)

            if has_serious_open:
                data["status"] = "RED"
                s_details = ", ".join(f"{s['name']} ({s.get('order_count', 0)} طلب | {s['amount']:,.2f} ر.س)" for s in serious_sessions)
                data["status_text"] = f"🚨 متأخر في الإقفال الرئيسي: الجلسة {s_details} لا تزال مفتوحة!"
                late_showrooms.append(data)
                if closed_unprinted:
                    data["main_session"] = closed_unprinted[0]
            elif closed_unprinted:
                main_sess = closed_unprinted[0]
                data["main_session"] = main_sess
                if data["secondary_sessions"]:
                    if is_month_end_alert:
                        data["status"] = "RED"
                        data["status_text"] = "🚨 تنبيه نهاية الشهر: يجب إقفال الجلسة الفرعية المعلقة فوراً!"
                        late_showrooms.append(data)
                    else:
                        data["status"] = "BLUE"
                        sec_details = ", ".join(f"{s['name']} ({s.get('order_count', 0)} طلب)" for s in data["secondary_sessions"])
                        data["status_text"] = f"🔵 جلسة فرعية شهرية مسموح ببقائها: {sec_details} (الجلسة الرئيسية جاهزة)"
                        secondary_open_showrooms.append(data)
                else:
                    data["status"] = "GREEN"
                    data["status_text"] = "جاهز تماماً للطباعة ✅"
                    ready_showrooms.append(data)
            else:
                # No closed unprinted session and no overdue open session
                if not open_sessions:
                    data["status"] = "GREEN"
                    data["status_text"] = "✅ تم إقفاله وطباعته سابقاً"
                    ready_showrooms.append(data)
                else:
                    data["status"] = "GREEN"
                    data["status_text"] = "جلسات مفتوحة حديثة (طبيعي) ✅"
                    ready_showrooms.append(data)

        return {
            "timestamp": now_ksa.strftime("%Y-%m-%d %H:%M:%S"),
            "cutoff_time_ksa": yesterday_5pm_ksa.strftime("%Y-%m-%d 17:00"),
            "is_month_end_alert": is_month_end_alert,
            "ready_count": len(ready_showrooms),
            "late_count": len(late_showrooms),
            "secondary_count": len(secondary_open_showrooms),
            "ready_showrooms": ready_showrooms,
            "late_showrooms": late_showrooms,
            "secondary_open_showrooms": secondary_open_showrooms,
            "all_showrooms": list(by_showroom.values())
        }

    def scan_morning_radar(self, target_date_str: str, allow_secondary_monthly: bool = True) -> Dict[str, Any]:
        """
        Morning Radar scan evaluating POS sessions for the target date.
        Distinguishes between main showroom sessions and permitted secondary monthly terminals.
        Returns the exact structure required by app.py.
        """
        if not self.is_connected() and not self.connect():
            raise ConnectionError("فشل الاتصال بسيرفر أودو. تأكد من إعدادات الاتصال وبيانات الدخول.")

        try:
            t_date = datetime.strptime(target_date_str, "%Y-%m-%d").date()
        except Exception:
            t_date = datetime.now(KSA_OFFSET).date()

        now_ksa = datetime.now(KSA_OFFSET)

        # 1. Fetch active POS configs
        pos_configs = self.search_read('pos.config', [('active', '=', True)], ['id', 'name'])
        config_names = {c['id']: c['name'] for c in pos_configs}

        # 2. Approved closings
        from database import SessionLocal, OldClosing
        old_closings = set()
        try:
            with SessionLocal() as db:
                old_closings = set(r[0] for r in db.query(OldClosing.closing_no).all())
        except Exception as e:
            print(f"[OdooClient] Note fetching old closings: {e}")

        # 3. Query window: 4 days before target date
        since_utc = (t_date - timedelta(days=4)).strftime("%Y-%m-%d 00:00:00")
        domain = [('start_at', '>=', since_utc)]
        fields = [
            'id', 'name', 'state', 'config_id', 'user_id',
            'start_at', 'stop_at', 'total_payments_amount', 'order_count'
        ]
        raw_sessions = self.search_read('pos.session', domain, fields)

        # Group by config_id
        by_config: Dict[int, List[Dict[str, Any]]] = {}
        for sess in raw_sessions:
            cid = sess['config_id'][0] if isinstance(sess.get('config_id'), (list, tuple)) else sess.get('config_id')
            if cid:
                by_config.setdefault(cid, []).append(sess)

        closed_ready_count = 0
        late_showrooms = []
        allowed_secondary_sessions = []

        for cfg_id, cfg_name in config_names.items():
            sess_list = by_config.get(cfg_id, [])

            # Closed unprinted sessions
            closed_unprinted = [
                s for s in sess_list
                if s.get('state') == 'closed' and s.get('name') and s['name'] not in old_closings
            ]
            if closed_unprinted:
                closed_ready_count += len(closed_unprinted)

            # Open sessions
            open_sessions = [s for s in sess_list if s.get('state') != 'closed']

            for osess in open_sessions:
                s_name = osess.get('name', '')
                s_start = osess.get('start_at', '')
                u_tuple = osess.get('user_id')
                opened_by = u_tuple[1] if isinstance(u_tuple, (list, tuple)) else "كاشير المعرض"
                order_cnt = int(osess.get('order_count', 0) or 0)
                amount = float(osess.get('total_payments_amount', 0.0) or 0.0)

                start_dt_ksa = None
                if s_start:
                    try:
                        dt_utc = datetime.strptime(s_start, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
                        start_dt_ksa = dt_utc.astimezone(KSA_OFFSET)
                    except Exception:
                        pass

                opened_at_str = start_dt_ksa.strftime("%Y-%m-%d %H:%M") if start_dt_ksa else (s_start or "غير محدد")

                # If target_date is before today (e.g. yesterday), any session opened today is today's shift (not late for yesterday)
                if start_dt_ksa and start_dt_ksa.date() == now_ksa.date() and t_date < now_ksa.date():
                    continue

                # If session started on today's morning and target_date is today, it's the active live working shift
                is_today_live_shift = False
                if start_dt_ksa and start_dt_ksa.date() == now_ksa.date() and t_date == now_ksa.date():
                    is_today_live_shift = True

                # Check if session is secondary monthly
                is_secondary = False
                if allow_secondary_monthly and not is_today_live_shift:
                    if closed_unprinted or (order_cnt <= 5 and amount < 1000.0):
                        is_secondary = True

                if is_secondary:
                    allowed_secondary_sessions.append({
                        "pos_name": cfg_name,
                        "session_name": s_name,
                        "opened_by": opened_by,
                        "opened_at": opened_at_str,
                        "status_text": f"جلسة فرعية شهرية ({order_cnt} طلب | {amount:,.2f} ر.س - مسموح ببقائها)"
                    })
                elif not is_today_live_shift:
                    late_showrooms.append({
                        "pos_name": cfg_name,
                        "session_name": s_name,
                        "opened_by": opened_by,
                        "opened_at": opened_at_str,
                        "status_text": f"🚨 متأخر في الإقفال: الجلسة لا تزال مفتوحة ({order_cnt} طلب | {amount:,.2f} ر.س)"
                    })

        total_pos_scanned = len(config_names) if config_names else len(by_config)
        stats = {
            "total_pos_scanned": total_pos_scanned,
            "closed_ready_count": closed_ready_count,
            "allowed_secondary_count": len(allowed_secondary_sessions),
            "late_count": len(late_showrooms)
        }

        return {
            "stats": stats,
            "ready_for_daily_report": len(late_showrooms) == 0,
            "late_showrooms": late_showrooms,
            "allowed_secondary_sessions": allowed_secondary_sessions
        }

    # =========================================================================
    # TAB 2: DIRECT SALES FETCHING FROM ODOO
    # =========================================================================
    def fetch_unprinted_closed_sales(self, old_closings: Set[str], days_back: int = 2) -> List[Dict[str, Any]]:
        """
        Directly fetches unprinted closed sessions from Odoo and calculates financial totals
        matching the exact Access Daily Sales structure.
        """
        now_utc = datetime.now(timezone.utc)
        since_utc_str = (now_utc - timedelta(days=days_back)).strftime("%Y-%m-%d 00:00:00")

        domain = [
            ('state', '=', 'closed'),
            ('stop_at', '>=', since_utc_str)
        ]
        fields = [
            'id', 'name', 'config_id', 'start_at', 'stop_at',
            'cash_register_balance_start', 'cash_register_balance_end_real',
            'cash_register_difference', 'total_payments_amount'
        ]

        sessions = self.search_read('pos.session', domain, fields)
        unprinted = [s for s in sessions if s['name'] not in old_closings]

        if not unprinted:
            return []

        session_ids = [s['id'] for s in unprinted]

        # Fetch orders for these sessions (including commission lines)
        order_domain = [('session_id', 'in', session_ids), ('state', 'in', ['paid', 'done', 'invoiced'])]
        order_fields = ['id', 'session_id', 'amount_total', 'amount_tax', 'amount_paid', 'amount_return', 'commission_line_ids']
        orders = self.search_read('pos.order', order_domain, order_fields)

        # Batch fetch salesmen commissions
        all_comm_ids = []
        for ord_item in orders:
            all_comm_ids.extend(ord_item.get('commission_line_ids', []))

        comm_map: Dict[int, float] = {}
        if all_comm_ids:
            try:
                c_records = self.search_read('commission.line', [('id', 'in', all_comm_ids)], ['id', 'commission'])
                comm_map = {c['id']: float(c.get('commission', 0.0) or 0.0) for c in c_records}
            except Exception as e:
                print(f"[OdooClient] Failed to fetch commission lines: {e}")

        # Fetch accounting closing payments (account.payment) for closed sessions
        # (This matches the exact Access & Pos Closing Daily Sales Report numbers down to the penny)
        acc_pay_domain = [('pos_session_id', 'in', session_ids)]
        acc_pay_fields = ['id', 'pos_session_id', 'amount', 'ref']
        acc_payments = []
        try:
            acc_payments = self.search_read('account.payment', acc_pay_domain, acc_pay_fields)
        except Exception as e:
            print(f"[OdooClient] Warning: Failed to fetch account.payment: {e}")

        # Protect against post-closing accounting modifications:
        # If an accountant altered account.payment amounts in Odoo after session closing,
        # retrieve the original amount signed at session closing from mail.tracking.value.
        orig_payment_amounts: Dict[int, float] = {}
        pay_ids = [p['id'] for p in acc_payments]
        if pay_ids:
            try:
                tracks = self.search_read(
                    'mail.tracking.value',
                    [
                        ('mail_message_id.model', '=', 'account.payment'),
                        ('mail_message_id.res_id', 'in', pay_ids),
                        ('field_id.name', 'in', ['amount', 'amount_signed', 'amount_company_currency_signed'])
                    ],
                    ['mail_message_id', 'old_value_float'],
                    limit=0
                )
                if tracks:
                    msg_ids = list({
                        t['mail_message_id'][0]
                        for t in tracks
                        if isinstance(t.get('mail_message_id'), (list, tuple))
                    })
                    msgs = self.search_read('mail.message', [('id', 'in', msg_ids)], ['res_id'])
                    msg_to_res = {m['id']: m['res_id'] for m in msgs}
                    # Earliest tracking is original showroom closing figure
                    tracks.sort(key=lambda x: x.get('id', 0))
                    for t in tracks:
                        m_tuple = t.get('mail_message_id')
                        m_id = m_tuple[0] if isinstance(m_tuple, (list, tuple)) else m_tuple
                        p_id = msg_to_res.get(m_id)
                        old_val = float(t.get('old_value_float') or 0.0)
                        if p_id and p_id not in orig_payment_amounts and old_val > 0.0:
                            orig_payment_amounts[p_id] = old_val
            except Exception as e:
                print(f"[OdooClient] Note on audit tracking check: {e}")

        # Also fetch pos.payment as a fallback for any sessions without account.payment
        payment_domain = [('session_id', 'in', session_ids)]
        payment_fields = ['id', 'session_id', 'payment_method_id', 'amount']
        payments = []
        try:
            payments = self.search_read('pos.payment', payment_domain, payment_fields)
        except Exception as e:
            print(f"[OdooClient] Warning: Failed to fetch pos.payment: {e}")

        pms = self.search_read('pos.payment.method', [], ['id', 'name', 'is_cash_count', 'is_payment_transfer'])
        pm_map = {m['id']: m for m in pms}

        # Fetch cash-out expenses from statement lines (e.g. POS/XXXX-out-مخرج)
        expenses_by_session: Dict[int, float] = {}
        try:
            st_domain = [('pos_session_id', 'in', session_ids), ('amount', '<', 0)]
            st_fields = ['id', 'pos_session_id', 'amount', 'payment_ref']
            st_lines = self.search_read('account.bank.statement.line', st_domain, st_fields)
            for l in st_lines:
                sid = l['pos_session_id'][0] if isinstance(l.get('pos_session_id'), (list, tuple)) else l.get('pos_session_id')
                ref_txt = (l.get('payment_ref') or '').lower()
                if 'out' in ref_txt or 'مخرج' in ref_txt:
                    expenses_by_session[sid] = expenses_by_session.get(sid, 0.0) + abs(float(l.get('amount', 0.0) or 0.0))
        except Exception as e:
            print(f"[OdooClient] Warning: Failed to fetch statement lines: {e}")

        # Aggregate orders per session
        orders_by_session: Dict[int, Dict[str, float]] = {}
        for ord_item in orders:
            s_id = ord_item['session_id'][0] if isinstance(ord_item['session_id'], (list, tuple)) else ord_item['session_id']
            if s_id not in orders_by_session:
                orders_by_session[s_id] = {'gross': 0.0, 'tax': 0.0, 'returns': 0.0, 'commissions': 0.0}
            tot = float(ord_item.get('amount_total', 0.0) or 0.0)
            tax = float(ord_item.get('amount_tax', 0.0) or 0.0)
            if tot < 0:
                orders_by_session[s_id]['returns'] += tot
            else:
                orders_by_session[s_id]['gross'] += tot
            orders_by_session[s_id]['tax'] += tax
            for cid in ord_item.get('commission_line_ids', []):
                orders_by_session[s_id]['commissions'] += comm_map.get(cid, 0.0)

        # Aggregate payments per session (Network vs Transfer) and capture accounting adjustments
        payments_by_session: Dict[int, Dict[str, float]] = {}
        adjustments_by_session: Dict[int, float] = {}
        sessions_with_acc_pay = set()

        for p in acc_payments:
            s_id = p['pos_session_id'][0] if isinstance(p.get('pos_session_id'), (list, tuple)) else p.get('pos_session_id')
            if s_id not in payments_by_session:
                payments_by_session[s_id] = {'cash': 0.0, 'credit': 0.0, 'trans': 0.0}
            sessions_with_acc_pay.add(s_id)
            ref_txt = (p.get('ref') or '').lower()

            p_id = p.get('id')
            current_amt = float(p.get('amount', 0.0) or 0.0)

            # Restore original showroom manager closing amount if modified post-closing
            if p_id in orig_payment_amounts:
                amt = orig_payment_amounts[p_id]
                # Accounting adjustment = (Current modified amount in Odoo) - (Original manager amount)
                # E.g. 1434.12 - 1745.72 = -311.60
                adj_delta = round(current_amt - amt, 2)
                adjustments_by_session[s_id] = adjustments_by_session.get(s_id, 0.0) + adj_delta
            else:
                amt = current_amt

            if 'تحويل' in ref_txt or 'trans' in ref_txt:
                payments_by_session[s_id]['trans'] += amt
            else:
                payments_by_session[s_id]['credit'] += amt

        # Fallback to pos.payment for sessions without account.payment
        for p in payments:
            s_id = p['session_id'][0] if isinstance(p.get('session_id'), (list, tuple)) else p.get('session_id')
            if s_id in sessions_with_acc_pay:
                continue  # Already accurately recorded via account.payment
            if s_id not in payments_by_session:
                payments_by_session[s_id] = {'cash': 0.0, 'credit': 0.0, 'trans': 0.0}
            pm_id = p['payment_method_id'][0] if isinstance(p.get('payment_method_id'), (list, tuple)) else p.get('payment_method_id')
            pm = pm_map.get(pm_id, {})
            p_name = pm.get('name', '').lower()
            amt = float(p.get('amount', 0.0) or 0.0)

            is_cash = pm.get('is_cash_count', False) or ('كاش' in p_name) or ('نقد' in p_name) or ('صندوق' in p_name)
            is_trans = pm.get('is_payment_transfer', False) or ('تحويل' in p_name) or ('trans' in p_name)

            if is_cash:
                payments_by_session[s_id]['cash'] += amt
            elif is_trans:
                payments_by_session[s_id]['trans'] += amt
            else:
                payments_by_session[s_id]['credit'] += amt

        # Build records
        records = []
        for s in unprinted:
            s_id = s['id']
            c_tuple = s.get('config_id')
            raw_desc = c_tuple[1] if isinstance(c_tuple, (list, tuple)) else "Showroom"
            closing_no = s.get('name', '')
            stop_at = s.get('stop_at', '')
            stop_at_ksa = stop_at
            if stop_at:
                try:
                    dt_utc = datetime.strptime(stop_at, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
                    stop_at_ksa = dt_utc.astimezone(KSA_OFFSET).strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass

            ord_stats = orders_by_session.get(s_id, {'gross': 0.0, 'tax': 0.0, 'returns': 0.0, 'commissions': 0.0})
            pay_stats = payments_by_session.get(s_id, {'cash': 0.0, 'credit': 0.0, 'trans': 0.0})

            net_total = ord_stats['gross'] + ord_stats['returns']  # Net total including returns
            tax = ord_stats['tax']
            net_sales = net_total - tax
            calc_total = net_total + ord_stats['returns']  # Matching Access formula: NetTotal + Returns

            open_cash = float(s.get('cash_register_balance_start', 0.0) or 0.0)
            actual_cash = float(s.get('cash_register_balance_end_real', 0.0) or 0.0)
            exp_val = expenses_by_session.get(s_id, 0.0)

            # Access Formula: Difference = (Network + Transfer + Cash Drawer + Expenses) - NetTotal
            diff = round((pay_stats['credit'] + pay_stats['trans'] + actual_cash + exp_val) - net_total, 2)

            records.append({
                "raw_desc": raw_desc,
                "closing_no": closing_no,
                "date_raw": stop_at_ksa,
                "total": calc_total,
                "returns": ord_stats['returns'],
                "tax": tax,
                "net_sales": net_sales,
                "net_total": net_total,
                "open": open_cash,
                "credit": pay_stats['credit'],
                "trans": pay_stats['trans'],
                "cash": actual_cash,
                "commissions": round(ord_stats['commissions']),
                "expenses": exp_val,
                "diff": diff,
                "acc_adjustment": round(adjustments_by_session.get(s_id, 0.0), 2)
            })

        return records

    def fetch_daily_pos_records(self, target_date_str: str, showroom_resolver: Any = None) -> List[Dict[str, Any]]:
        """
        Fetches closed pos.sessions directly from Odoo via XML-RPC.
        Resolves showroom metadata, formats dates, and prepares records for the daily sales editor.
        """
        from database import SessionLocal, OldClosing
        from processor import parse_date_value

        old_closings = set()
        try:
            with SessionLocal() as db:
                old_closings = set(r[0] for r in db.query(OldClosing.closing_no).all())
        except Exception as e:
            print(f"[OdooClient] Note fetching old closings: {e}")

        raw_sales = self.fetch_unprinted_closed_sales(old_closings=old_closings, days_back=5)

        formatted_records = []
        for idx, r in enumerate(raw_sales, start=1):
            raw_desc = r.get("raw_desc", "")

            sh = showroom_resolver.resolve(raw_desc) if showroom_resolver else None
            short_name = sh.Short if sh else raw_desc
            city_no = sh.CityNo if sh and sh.CityNo else "999"
            city_name = "غير محدد"
            branch_name = "عام"
            if sh:
                try:
                    if sh.city:
                        city_name = sh.city.CityName
                        if sh.city.branch:
                            branch_name = sh.city.branch.BranchName
                except Exception:
                    pass

            date_raw = r.get("date_raw", "")
            date_str, closing_hour = parse_date_value(date_raw)

            full_rec = {
                "selected": True,
                "is_old": False,
                "seq": idx,
                "raw_desc": raw_desc,
                "showroom": short_name,
                "city_no": city_no,
                "city_name": city_name,
                "branch_name": branch_name,
                "closing_no": r.get("closing_no", ""),
                "date": date_str,
                "closing_hour": closing_hour,
                "total": r.get("total", 0.0),
                "returns": r.get("returns", 0.0),
                "tax": r.get("tax", 0.0),
                "net_sales": r.get("net_sales", 0.0),
                "net_total": r.get("net_total", 0.0),
                "open": r.get("open", 0.0),
                "credit": r.get("credit", 0.0),
                "trans": r.get("trans", 0.0),
                "cash": r.get("cash", 0.0),
                "commissions": r.get("commissions", 0),
                "expenses": r.get("expenses", 0.0),
                "diff": r.get("diff", 0.0),
                "acc_adjustment": r.get("acc_adjustment", 0.0)
            }
            formatted_records.append(full_rec)

        # Sort: ORDER BY CityNo ASC, Showroom ASC, date ASC
        formatted_records.sort(key=lambda x: (x["city_no"], x["showroom"], x["date"]))
        for i, rec in enumerate(formatted_records, start=1):
            rec["seq"] = i

        return formatted_records

    # =========================================================================
    # TAB 4: PERIODIC & COMPARATIVE EXECUTIVE SALES REPORTS
    # =========================================================================
    def fetch_period_sales(
        self,
        start_date: str,
        end_date: str,
        title: str = "تقرير مبيعات المعارض",
        period_type: str = "custom",
        showroom_resolver: Any = None
    ) -> Dict[str, Any]:
        """
        Fetches and aggregates closed sales per showroom for any given date range (YYYY-MM-DD),
        matching the official executive reporting format (sorted by Net descending).
        """
        import re

        s_y, s_m, s_d = [int(x) for x in start_date.split('-')]
        e_y, e_m, e_d = [int(x) for x in end_date.split('-')]

        start_dt_utc = datetime(s_y, s_m, s_d, 0, 0, 0, tzinfo=KSA_OFFSET).astimezone(timezone.utc)
        end_dt_utc = datetime(e_y, e_m, e_d, 23, 59, 59, tzinfo=KSA_OFFSET).astimezone(timezone.utc)
        start_str = start_dt_utc.strftime("%Y-%m-%d %H:%M:%S")
        end_str = end_dt_utc.strftime("%Y-%m-%d %H:%M:%S")

        if not self.is_connected() and not self.connect():
            raise ConnectionError("Could not connect to Odoo server. Check credentials.")

        domain = [
            ('state', 'in', ['paid', 'done', 'invoiced']),
            ('date_order', '>=', start_str),
            ('date_order', '<=', end_str)
        ]

        # Single high-speed read_group query on pos.order
        res = self._models.execute_kw(
            self.db, self.uid, self.password,
            'pos.order', 'read_group',
            [domain],
            {
                'fields': ['config_id', 'amount_total:sum', 'amount_tax:sum'],
                'groupby': ['config_id'],
                'lazy': False
            }
        )

        by_code: Dict[str, Dict[str, Any]] = {}

        for r in res:
            cfg = r.get('config_id')
            name = cfg[1] if isinstance(cfg, (list, tuple)) else str(cfg)
            tot = float(r.get('amount_total', 0.0) or 0.0)
            tax = float(r.get('amount_tax', 0.0) or 0.0)

            # Match 3-digit showroom code (101 to 199)
            m = re.findall(r'\b(1\d\d)\b', name)
            code = m[0] if m else None

            sh = None
            if showroom_resolver:
                if code and f"NO_{code}" in showroom_resolver.mapping:
                    sh = showroom_resolver.mapping[f"NO_{code}"]
                else:
                    sh = showroom_resolver.resolve(name)
                if sh and not code:
                    code = sh.No

            if not code:
                continue

            short = sh.Short if sh else name
            # Special check for Al-Bawadi MS (122) vs Al-Bawadi (112)
            if "ms" in name.lower() and "122" in name:
                sh_display = "معرض البوادي MS"
            elif "معرض" in short:
                sh_display = short
            else:
                sh_display = f"معرض {short}"

            if code not in by_code:
                by_code[code] = {
                    'code': code,
                    'showroom': sh_display,
                    'total': 0.0,
                    'tax': 0.0
                }
            by_code[code]['total'] += tot
            by_code[code]['tax'] += tax

        rows = []
        for code, v in by_code.items():
            tot = round(v['total'], 2)
            tax = round(v['tax'], 2)
            net = round(tot - tax, 2)
            rows.append({
                'code': code,
                'showroom': v['showroom'],
                'total': tot,
                'tax': tax,
                'net': net
            })

        # Sort strictly by Net sales descending (من الأعلى إلى الأقل)
        rows.sort(key=lambda x: x['net'], reverse=True)

        # Assign rank 1 to N
        for idx, r in enumerate(rows, start=1):
            r['seq'] = idx

        grand_totals = {
            'total': round(sum(r['total'] for r in rows), 2),
            'tax': round(sum(r['tax'] for r in rows), 2),
            'net': round(sum(r['net'] for r in rows), 2),
            'showrooms_count': len(rows)
        }

        return {
            'title': title,
            'period_type': period_type,
            'start_date': start_date,
            'end_date': end_date,
            'rows': rows,
            'grand_totals': grand_totals
        }

    def fetch_monthly_sales(self, year: int, month: int, showroom_resolver: Any = None) -> Dict[str, Any]:
        """
        Fetches and aggregates closed sales per showroom for a specific calendar month,
        matching the official monthly executive report exactly (sorted by Net descending).
        """
        from calendar import monthrange
        _, last_day = monthrange(year, month)
        start_date = f"{year}-{month:02d}-01"
        end_date = f"{year}-{month:02d}-{last_day:02d}"
        title = f"مبيعات معارض العسلي لشهر {month} - {year}"

        data = self.fetch_period_sales(
            start_date=start_date,
            end_date=end_date,
            title=title,
            period_type="monthly",
            showroom_resolver=showroom_resolver
        )
        data['year'] = year
        data['month'] = month
        return data

    def fetch_comparative_sales(
        self,
        current_period: Dict[str, Any],
        previous_period: Dict[str, Any],
        title: str = "تقرير مقارنة مبيعات المعارض",
        showroom_resolver: Any = None
    ) -> Dict[str, Any]:
        """
        Compares two distinct periods (e.g. Month vs Previous Month, or YoY),
        calculating differences and growth percentages per showroom and company-wide.
        """
        curr_data = self.fetch_period_sales(
            start_date=current_period['start_date'],
            end_date=current_period['end_date'],
            title=current_period.get('label', 'الفترة الحالية'),
            showroom_resolver=showroom_resolver
        )
        prev_data = self.fetch_period_sales(
            start_date=previous_period['start_date'],
            end_date=previous_period['end_date'],
            title=previous_period.get('label', 'فترة المقارنة'),
            showroom_resolver=showroom_resolver
        )

        curr_map = {r['code']: r for r in curr_data['rows']}
        prev_map = {r['code']: r for r in prev_data['rows']}

        all_codes = set(curr_map.keys()) | set(prev_map.keys())
        rows = []

        for code in all_codes:
            curr_r = curr_map.get(code)
            prev_r = prev_map.get(code)

            sh_name = (curr_r['showroom'] if curr_r else None) or (prev_r['showroom'] if prev_r else f"معرض {code}")
            curr_net = curr_r['net'] if curr_r else 0.0
            prev_net = prev_r['net'] if prev_r else 0.0

            diff = round(curr_net - prev_net, 2)
            if prev_net > 0:
                growth_pct = round((diff / prev_net) * 100.0, 2)
            else:
                growth_pct = 100.0 if curr_net > 0 else 0.0

            rows.append({
                'code': code,
                'showroom': sh_name,
                'current_net': curr_net,
                'previous_net': prev_net,
                'diff': diff,
                'growth_pct': growth_pct
            })

        # Sort by current Net sales descending
        rows.sort(key=lambda x: x['current_net'], reverse=True)

        for idx, r in enumerate(rows, start=1):
            r['seq'] = idx

        curr_tot = round(sum(r['current_net'] for r in rows), 2)
        prev_tot = round(sum(r['previous_net'] for r in rows), 2)
        diff_tot = round(curr_tot - prev_tot, 2)
        growth_tot = round((diff_tot / prev_tot) * 100.0, 2) if prev_tot > 0 else 0.0

        grand_totals = {
            'current_net': curr_tot,
            'previous_net': prev_tot,
            'diff': diff_tot,
            'growth_pct': growth_tot,
            'showrooms_count': len(rows)
        }

        return {
            'title': title,
            'current_label': current_period.get('label', 'الفترة الحالية'),
            'previous_label': previous_period.get('label', 'فترة المقارنة'),
            'current_range': f"{current_period['start_date']} إلى {current_period['end_date']}",
            'previous_range': f"{previous_period['start_date']} إلى {previous_period['end_date']}",
            'rows': rows,
            'grand_totals': grand_totals
        }
