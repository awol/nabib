#!/usr/bin/env python
'''
    A module to parse the internet banking website of the National Australia Bank and produce QIF files for the accounts
    in the given user account.
'''
import sys
import time
import decimal
import argparse
import datetime
import math
import re
import textwrap
import pathlib
import os
import traceback
import logging
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException, WebDriverException

APP_DETAILS = {
    'nab-app-id': 'cf87dc5d-0245-4eff-8d99-37f2da85bf44'
}
DECIMAL_CONTEXT = decimal.Context(prec=2, rounding=decimal.ROUND_HALF_DOWN)
decimal.setcontext(DECIMAL_CONTEXT)
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(format=FORMAT)
logger = logging.getLogger(__name__)

def print_delay(length, period=1):
    '''\
       Delay the execution of the program for a number of seconds, sleeping for a configurable
       number of seconds repeatedly until the total delay has elapsed
       '''
    try:
        print 'Sleeping for ',
        for i in range(length, 0, -period):
            print '{}, '.format(i),
            sys.stdout.flush()
            time.sleep(period)
    except KeyboardInterrupt:
        pass
    print ''


class UnexpectedPageTitle(Exception):
    '''We got the wrong page'''
    pass


def assert_title(driver, expected_title):
    '''Raise an exception if the title of the current page does not match the expected title'''
    if driver.title != expected_title:
        raise UnexpectedPageTitle('Expected page "{}", got "{}"'.format(expected_title, driver.title))


class MonthDate(datetime.date):
    '''A class of date that has a number of month end related functions'''
    def __iter__(self):
        for attr in ['year', 'month', 'day']:
            yield getattr(self, attr)

    @property
    def month_start(self):
        '''return the start of the month for the date represented by this instance'''
        return MonthDate(*(tuple(self)[:2] + (1, )))

    @property
    def month_end(self):
        '''return the end of the month for the date represented by this instance'''
        return self.month_next.month_end_prev

    @property
    def month_end_prev(self):
        '''return the end of the previous month for the date represented by this instance'''
        return MonthDate.fromordinal((self.month_start - datetime.timedelta(days=1)).toordinal())

    @property
    def month_prev(self):
        '''return the start of the previous month for the date represented by this instance'''
        return MonthDate.fromordinal((self.month_start - datetime.timedelta(days=1)).toordinal()).month_start

    @property
    def month_next(self):
        '''return the start of the next month for the date represented by this instance'''
        return MonthDate.fromordinal((self.month_start + datetime.timedelta(days=31)).toordinal()).month_start

    @classmethod
    def strptime(cls, value, fmt_spec):
        '''return the instance of the date defined by the value and the format specification passed into the function'''
        timestamp = datetime.datetime.strptime(value, fmt_spec)
        use_date = timestamp.date()
        return MonthDate.fromordinal(use_date.toordinal())


class NABNumber(decimal.Decimal):
    '''A decimal number that supports the kind of formatting NAB internet banking uses'''
    def __new__(cls, value, **kwargs):
        '''meth_doc'''
        if isinstance(value, decimal.Decimal):
            return decimal.Decimal.__new__(cls, value)
        value = re.sub(r'[+,$]', r'', ''.join(['{}'.format(x).strip() for x in (value or 'NaN').split('\n')])).strip()
        number, dorc = (value.split(' ') + [''])[:2]
        number = number.replace('-minus', '-')
        if value in ['N/A', '']:
            convert_value = 'NaN'
        else:
            convert_value = '{}{}'.format('-' if dorc == 'DR' else '', number)
        logger.debug('v:%s: n:%s: dc:%s: cv:%s:', value, number, dorc, convert_value)
        return decimal.Decimal.__new__(cls, convert_value, **kwargs)

    def __str__(self):
        '''Return a string representation of the number, showing two decomal places. Essentially a currency amount.'''
        string_rep = '{:6.2f}'.format(self)
        return string_rep

    def __add__(self, rhs):
        '''meth_doc'''
        return NABNumber(decimal.Decimal.__add__(self, rhs))

    def __iadd__(self, rhs):
        '''meth_doc'''
        return NABNumber(decimal.Decimal.__add__(self, rhs))


class Transaction(object):
    '''A transaction from the internet banking site'''
    def __init__(self, date, trans_type, reference, payee, memo, amount=None, category='Unspecified', balance=None):
        '''meth_doc'''
        self.date = date
        self.type = trans_type
        self.reference = reference
        self.payee = payee
        self.memo = memo
        self.amount = amount
        self.category = category
        self.balance = balance

    # Output the transaction in QIF format
    @property
    def qif(self):
        '''meth_doc'''
        qif_string = textwrap.dedent('''\
            D{t.date}
            T{a}
            M{t.memo}
            N{t.type}
            P{t.payee}
            L{t.category}
            ^
            '''.format(t=self, a=str(self.amount).strip()))
        return qif_string

    def __str__(self):
        return '|'.join(['{}'.format(getattr(self, x)) for x in ['date', 'amount', 'memo', 'type', 'payee', 'category']])

    @property
    def csv(self):
        '''meth_doc'''
        csv_string = textwrap.dedent('''\
            {t.date},\
            {a},\
            {t.memo},\
            {t.type},\
            {t.payee},\
            {t.category}\
            '''.format(t=self, a=str(self.amount).strip()))
        return csv_string


def wrap_for_unexpected_alert(driver, action, log_message=True):
    '''wrap the action (a callable) in a try block that will catch a browser alert'''
    result = None
    try:
        result = action()
        severity, code, msg_text = ('', '', '')
    except UnexpectedAlertPresentException:
        alert = driver.switch_to.alert
        severity, code, msg_text = (lambda x, y, z: [x.lower().strip(), y.lower().strip(':').strip(), z.strip()])(*((['info', 'unknown'] + alert.text.split(' ', 2))[-3:]))
        if log_message:
            logger.info('%s: [%s] %s', severity, code, msg_text)
        alert.dismiss()
    return result, severity, code, msg_text


def wait_spinner(driver, timeout=10):
    '''meth_doc'''
    for index in range(timeout):
        spinner_count = len(driver.find_elements_by_xpath("//div[contains(@class, 'transaction-history-spinner-container')]"))
        if spinner_count > 0:
            time.sleep(index)
        else:
            break

class Account(object):
    '''An account from the internet banking site'''
    # The mappings between NAB Account Types and the types of account defined in the QIF specification.
    acct_type_map = {
        'DDA': 'Bank', 'SDA': 'Bank', 'VCD': 'CCard',
        'transaction-account': 'Bank', 'credit-card': 'CCard'
    }

    def __init__(self, trans_type, number, nick_name=None, opening_balance=0.0, available_balance=None, at_date=datetime.date.today()):
        '''meth_doc'''
        self.var_names = [
            'date',
            'details',
            'debit',
            'credit',
            'balance',
        ]
        self.type = trans_type.strip("'")
        self.bsb = None
        if re.match(r'BSB:.*Acct No:.*', number):
            self.bsb, self.number = re.sub(r'(BSB: *|Acc: *|  )', r'', number).strip().split(' ', 1)
        elif re.match(r'Card ending .*', number):
            self.bsb = None
            self.number = next(iter(re.sub(r'(Card ending *|  )', r'', number).strip().split(' ', 1)), number)
        else:
            self.number = number
        if nick_name:
            self.nick_name = nick_name
        else:
            self.nick_name = self.identifier
        self.current_balance = opening_balance
        if available_balance:
            self.available_balance = available_balance
        else:
            self.available_balance = opening_balance
        self.at_date = at_date
        self.closing_balance = NABNumber('0.00')
        self.closing_balance_date = datetime.date.today()
        self.transactions = []
        self._payee_category_map = None

    # This function will generate the id used internally by NAB to identify the account within the internet banking application
    @property
    def identifier(self):
        '''meth_doc'''
        return re.sub(r'[^0-9]', r'', self.number)

    @property
    def payee_category_map(self):
        '''meth_doc'''
        if getattr(self, '_payee_category_map', None) is None:
            self._payee_category_map = {}
            for file_name in ["PayeeCategories-" + self.nick_name + ".txt", "PayeeCategories.txt"]:
                file_path = os.path.join(os.path.dirname(sys.argv[0]), file_name)
                if os.path.isfile(file_path):
                    with open(file_path) as in_file:
                        for line in in_file.readlines():
                            if not line.strip():
                                continue
                            payee, category = line.strip().split('|')
                            self._payee_category_map[payee] = category
        return self._payee_category_map

    @property
    def qif(self):
        '''meth_doc'''
        qif_string = textwrap.dedent('''\
            !Account
            N{}{}{n}
            T{t}
            '''.format(*([self.bsb, ' '] if self.bsb is not None else ['', '']), n=self.number, t=self.acct_type_map.get(self.type, 'Bank')))
        self.closing_balance = self.current_balance
        self.closing_balance_date = self.at_date
        trans_qif_string = ''
        if self.transactions:
            for trans in self.transactions:
                trans_qif_string += trans.qif
                self.closing_balance = trans.balance
                self.closing_balance_date = trans.date
        qif_string += textwrap.dedent('''\
            ${cb}
            /{cbd}
            ^
            !Type:{type}
            '''.format(cbd=self.closing_balance_date, cb=('0.00' if math.isnan(self.closing_balance) else self.closing_balance), type=self.acct_type_map.get(self.type, 'Bank')))
        qif_string += trans_qif_string
        return qif_string

    def filter_transactions(self, driver, start_date, end_date):
        '''meth_doc'''
        try_duration = 120
        sleep_time = 1
        start_time = datetime.datetime.now()
        while int((datetime.datetime.now() - start_time).total_seconds()) < try_duration:
            try:
                logger.debug('try find filter button:{}'.format(start_time))
                driver.find_element_by_xpath("//button[@data-test-id='filter-button' and @aria-expanded='false']").click()
                break
            except (NoSuchElementException, TypeError, Exception) as excp:
                logger.debug('Exception:%s', excp)
                time.sleep(sleep_time)
                sleep_time += 2
        transaction_form = driver.find_element_by_xpath("//form[@name='filterForm']")
        transaction_form.find_element_by_xpath("//div[@id='accountSelect']//i[contains(@class, 'ion-ios-arrow-down')]").click()
        transaction_form.find_element_by_xpath("//div[@id='accountSelect']//ul[@id='ui-select-choices-0']/li//p[contains(@class, 'account-nickname') and text()='{}']".format(self.nick_name)).click()

        btn_elem = transaction_form.find_element_by_xpath("//div[@id='input-transaction-period']//i[contains(@class, 'ion-ios-arrow-down')]")
        logger.debug('"Got btn_elem:{}:'.format(btn_elem, type(btn_elem)))
        btn_elem.click()
        time.sleep(1)
        transaction_form.find_element_by_xpath("//div[@id='input-transaction-period']//ul//li//span[text()='Custom date range']").click()
        time.sleep(1)
        if start_date:
            elem = transaction_form.find_element_by_xpath("//input[@name='startDate']")
            elem.clear()
            elem.send_keys(start_date.strftime('%d/%m/%y'))
        if end_date:
            elem = transaction_form.find_element_by_xpath("//input[@name='endDate']")
            elem.clear()
            elem.send_keys(end_date.strftime('%d/%m/%y'))

        transaction_form.find_element_by_xpath("//div[@id='input-page-size']//i[contains(@class, 'ion-ios-arrow-down')]").click()
        transaction_form.find_element_by_xpath("//div[@id='input-page-size']//ul//li//span[text()='200']").click()

        transaction_form.find_element_by_xpath("//button[@id='displayBtn']").click()
        wait_spinner(driver)

    def download_transactions(self, driver, start_date, end_date):
        '''meth_doc'''
        if driver.title.lower() != 'Transaction History'.lower():
            driver.execute_script("sendMenuRequest('transactionHistorySelectAccount.ctl');")
        try:
            #driver.find_element_by_xpath("//div[@id='filter-button-section']//button[@data-test-id='filter-button']").click()
            driver.find_element_by_xpath("//button[@data-test-id='filter-button']").click()
        except NoSuchElementException:
            pass
        except Exception as excp:
            logger.info('DBG: excp:{}:'.format(excp))

        self.filter_transactions(driver, start_date, end_date)
        if wrap_for_unexpected_alert(driver, lambda: driver.find_elements_by_xpath("//table[@id='transactionHistoryTable']"))[1:3] == (u'error', u'302033'):
            self.var_names.pop()
        assert_title(driver, 'Transaction History')
        transactions = []
        data_buttons = [None] + driver.find_elements_by_xpath("//ul[contains(@class, 'pagination')]//li/button[contains(@class, 'btn-pagination')]")
        #transaction_rows = driver.find_elements_by_xpath("//div[@id='transactions']//app-component/ib-transactions//table[contains(@class, 'transaction-history-table')]/tbody/tr[not(contains(@class, 'hidden'))]")
        transaction_rows = driver.find_elements_by_xpath("//div[@id='transactions']//app-component/ib-transactions//table[contains(@class, 'transaction-history-table')]/tbody/tr[not(contains(@class, 'hidden'))]")
        if len(transaction_rows) <= 0:
            return NABNumber('0.00')
        logger.info('      Processing %s transactions for account "%s" (%s pages)', len(transaction_rows), self.nick_name, len(data_buttons))
        #for data_button in data_buttons:
            #if data_button and not re.match(r'.*btn-pagination-current.*', data_button.get_attribute("class", '')):
                #data_button.click()
                #wait_spinner(driver)
        #for page_index in range(len(data_buttons)):
        time.sleep(3)
        for index, transaction_row in enumerate(transaction_rows):
            transactions.append(self.process_row(index, transaction_row))
        self.transactions = list(reversed(transactions))
        logger.debug('    Finished %s transactions for account "%s"', len(self.transactions), self.nick_name)
        self.closing_balance = getattr(next(iter(self.transactions[-1:]), None), 'balance', NABNumber('0.00'))
        self.closing_balance_date = getattr(next(iter(self.transactions[-1:]), None), 'date', end_date)
        return self.closing_balance

    def process_row(self, index, transaction_row):
        '''meth_doc'''
        logger.debug('index:%s: tr:%s: text:%s:', index, transaction_row, transaction_row.text)
        cells = transaction_row.find_elements_by_xpath("./td")
        logger.debug('DBG: c:%s:', cells)
        if len(cells) < len(self.var_names):
            logger.info('skipping row, lc:%s: lv:%s: t:%s: c:%s:', len(cells), len(self.var_names), transaction_row.text, [x.text for x in cells])
            return
        values = {x: transaction_row.find_element_by_xpath('./td[{}]'.format(i)) for i, x in enumerate(self.var_names, 1)}
        values['date'] = getattr(values['date'], 'text', None)
        values['location'], values['memo'] = list(reversed(re.split(r'   *', getattr(values['details'].find_element_by_xpath('.//div[@ng-bind-html="transaction.narrative"]'), 'text', '').strip(), 1))) + ['']
        if values['memo'] == '':
            values['memo'] = values['location']
            values['location'] = ''
        else:
            if isinstance(values['memo'], (list, tuple)):
                values['memo'] = ' '.join(reversed(values['memo']))
        values['trans_type'] = getattr(values['details'].find_element_by_xpath('.//span[@ng-bind-html="transaction.transCodeText"]'), 'text', '')
        values['debit'] = NABNumber((lambda x: '{}{}'.format('-' if x else '', x or 'NaN'))(getattr(next(iter(values['debit'].find_elements_by_xpath('.//span[@ng-bind-html="transaction.debitFormatted"]')), None), 'text', None)))
        values['credit'] = NABNumber(getattr(next(iter(values['credit'].find_elements_by_xpath('.//span[@ng-bind-html="transaction.creditFormatted"]')), None), 'text', None))
        prefix_text = ''
        try:
            prefix_text = getattr(values['balance'].find_element_by_xpath('./number_balance'), 'text', '')
        except NoSuchElementException:
            pass
        balance_text = re.sub(prefix_text, r'', getattr(values['balance'], 'text', 'NaN'))
        logger.debug('prefix:%s: balance:%s: new:%s:', prefix_text, getattr(values['balance'], 'text', 'NaN'), balance_text)
        values['balance'] = NABNumber(balance_text)
        logger.debug('1 values:%s:', values)
        amt = NABNumber('0.00')
        if not math.isnan(values['credit']):
            amt = values['credit']
        if not math.isnan(values['debit']):
            amt = values['debit']
        values['payee'] = values['memo'].replace(r'^.*[0-9][0-9]:[0-9][0-9] ', '').replace(r'^INTERNET BPAY *', '').replace(r'^INTERNET TRANSFER *', '').replace(r'^FEES *', '') if values['memo'] else None
        done_category = False
        for field_to_map in [values[x] for x in ['payee', 'memo']]:
            for re_key, value in self.payee_category_map.items():
                if re.match(r'^.*{}.*'.format(re_key), field_to_map or '', re.IGNORECASE):
                    values['category'] = value
                    done_category = True
                    break
            if done_category:
                break
        logger.debug('2 values:%s', values)
        return Transaction(MonthDate.strptime(values['date'], '%d %b %y'), *[values.get(x, None) for x in ['trans_type', 'reference', 'payee', 'memo']], amount=amt, category=values.get('category', 'Unspecified'), balance=values['balance'])

    def generate_qif(self, driver, start_date=MonthDate(*(datetime.date.today().timetuple()[:2] + (1, ))), end_date=MonthDate(*(datetime.date.today().timetuple()[:3])), **kwargs):
        '''meth_doc'''
        output_file = kwargs.pop('output_file', '{}.qif'.format(self.nick_name))
        logger.info('   Generating QIF for "%s" account (%s %s) in file "%s" from %s to %s', self.nick_name, self.bsb, self.number, os.path.realpath(output_file), start_date, end_date)
        closing_balance = self.download_transactions(driver, start_date, end_date, **kwargs)
        with open(output_file, 'w') as out_fh:
            print>>out_fh, self.qif
        return closing_balance


class SavedPages(object):
    '''A class to load pages from disk rather than from the website'''
    page_file_map = {
        'acctInfo_acctBal.ctl': 'Account summary',
        'transactionHistoryValidate.ctl': 'Transaction history',
        'transactionHistoryGetSettings.ctl': 'Transaction history',
        'transactionHistoryDisplay.ctl': 'Transaction history',
    }
    def __init__(self, test_key=None, page_path=None):
        '''meth_doc'''
        self.path = os.path.realpath(page_path or os.path.join(os.path.dirname(sys.argv[0]), 'pages'))
        self.test_key = test_key

    def get_page(self, name, test_key=None):
        '''meth_doc'''
        use_key = test_key or self.test_key
        test_keys = ([] if use_key is None else [use_key]) + ['']
        for try_key in test_keys:
            use_path = os.path.join(self.path, try_key)
            for ext in ['', '.html', 'htm']:
                use_file = os.path.join(use_path, '{}{}'.format(name, ext))
                if os.path.isfile(use_file):
                    return use_file
                use_file = os.path.join(use_path, '{}{}'.format(self.page_file_map.get(name, name), ext))
                if os.path.isfile(use_file):
                    return use_file
        return ''

    def save_page(self):
        '''meth_doc'''
        pass


def get_page(driver, name, **kwargs):
    '''a function to get a page from the given driver or file defined in the kwargs'''
    app_root = 'https://ib.nab.com.au/nabib'
    saved_pages = kwargs.pop('saved_pages', None)
    test_mode = kwargs.pop('test_mode', False)
    if saved_pages is not None and test_mode:
        page_path = saved_pages.get_page(name, **kwargs)
        return driver.get(pathlib.Path(page_path).as_uri())
    return driver.get('{}/{}'.format(app_root, name))


def get_accounts(driver, end_date, **kwargs):
    '''Get the accounts defined in the NAB internet banking account currently loaded in the passed web driver'''
    _accounts = {}
    logger.debug('get_accounts driver.title:{}:'.format(driver.title))
    if driver.title != 'Account summary':
        wrap_for_unexpected_alert(driver, lambda: get_page(driver, 'acctInfo_acctBal.ctl', **kwargs))
    assert_title(driver, 'Account summary')
    for account_row in driver.find_elements_by_xpath("//table[contains(@class, 'traditional-account-table')][1]/tbody/tr"):
        logger.debug('Found account row')
        var_selectors = {
            'nick_name': './td[1]//div[contains(@class, "account-nickname")]',
            'trans_type': './/nui-icon',
            'number': './td[1]//div[contains(@class, "account-number")]',
            'current': './td[2]',
            'available': './td[3]',
        }
        values = {x: getattr(account_row.find_element_by_xpath(y), 'text', None) for x, y in var_selectors.items()}
        logger.debug('Found account row values:{}:'.format(values))
        ttv= account_row.find_element_by_xpath(var_selectors['trans_type'])
        logger.debug('Found account row trans_type dir:{}: vars:{}:'.format(dir(ttv), vars(ttv)))
        trans_type = account_row.find_element_by_xpath(var_selectors['trans_type']).get_attribute('name') or u'transaction-account'
        logger.debug('Found account row attribute trans_type:{}: values:{}:'.format(trans_type, values))
        trans_type = account_row.find_element_by_xpath(var_selectors['trans_type']).get_property('name') or u'transaction-account'
        logger.debug('Found account row property trans_type:{}: values:{}:'.format(trans_type, values))
        _accounts[values['nick_name']] = Account(trans_type, values['number'], values['nick_name'], available_balance=NABNumber(values['available']), at_date=end_date)
    logger.debug('get_accounts accounts:{}:'.format(_accounts))
    return _accounts


def connect(driver, userid, password):
    '''Connect to the NAB internet banking account for the user and password'''
    logger.info('Launching and connecting (may take half a minute or so)')
    get_page(driver, 'index.jsp')
    elem = driver.find_element_by_name('userid')
    try:
        elem.clear()
    except WebDriverException:
        pass
    elem.send_keys(userid)
    elem = driver.find_element_by_name('password')
    elem.clear()
    elem.send_keys(password)
    time.sleep(2)
    elem.send_keys(Keys.RETURN)
    logger.debug('DBG sleep 10')
    time.sleep(10)
    logger.debug('DBG slept 10')
    assert_title(driver, 'Account summary')
    logger.debug('Connected')


def get_command_line_options():
    '''parse the command line options'''
    parser = argparse.ArgumentParser(description=textwrap.dedent(main.__doc__), formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    date_group = parser.add_mutually_exclusive_group()
    date_group.add_argument('--last-month', action='store_true', help='Get all the transactions from the last calendar month')
    date_group.add_argument('--this-month', action='store_true', help='Get all the transactions from the start of this calendar month until now')
    date_group.add_argument('--start-date', metavar='YYYYMMDD', type=lambda d: MonthDate.strptime(d, '%Y%m%d'), help='Get the transactions starting from this date (inclusive)', default=MonthDate.today().month_prev)
    parser.add_argument('--end-date', metavar='YYYYMMDD', type=lambda d: MonthDate.strptime(d, '%Y%m%d'), help='Get the transactions up until this date (inclusive)', default=None)
    parser.add_argument('--saved-pages', metavar='PATH', help='Use the pages downloaded into PATH when running in test mode')
    parser.add_argument('--test-mode', action='store_true', help='Test this application with saved pages rather than connecting to the actua Internet Banking app', default=None)
    parser.add_argument('--driver-keep', action='store', choices=['never', 'exception', 'always'], help='Keep the running browser after the program exits', default='never')
    parser.add_argument('--log-level', action='store', choices=[x for x in logging._levelNames if isinstance(x, str)], help='Set the log level', default='INFO')
    parser.add_argument('user', metavar='USER', help='The NAB internet banking customer number')
    parser.add_argument('password', metavar='PASSWORD', help='The password of the user provided')
    parser.add_argument('account', action='store', nargs='*', metavar='ACC', help='An account for which to generate a QIF file', default=[])
    cl_opts = parser.parse_args()
    today = MonthDate.today()
    if cl_opts.this_month:
        cl_opts.start_date = today.month_start
    elif cl_opts.last_month:
        cl_opts.start_date = today.month_prev
    if cl_opts.end_date is None:
        cl_opts.end_date = cl_opts.start_date.month_end
    return cl_opts


def main():
    '''\
       This program will connect to the NAB Internet banking site
       and download all the transactions available for the accounts in the internet banking account
       associated with the banking customer number provided on the command line.

       By default it will return the transactions from the whole month previous to the month of the
       current date at the time of invocation.

       This program uses selenium to interact with the Google Chrome browser to download transactional
       data.
       '''
    start_time = datetime.datetime.now()
    options = get_command_line_options()
    logger.setLevel(getattr(logging, options.log_level))
    web_driver = webdriver.Chrome()
    logger.debug('Started %s web driver: driver_keep:%s', web_driver.__class__.__name__, options.driver_keep)
    keep_driver = (options.driver_keep == 'always')
    try:
        options_kwargs = {}
        if options.saved_pages is not None or options.test_mode:
            options_kwargs['saved_pages'] = SavedPages(page_path=options.saved_pages)
        options_kwargs['test_mode'] = options.test_mode
        connect(web_driver, options.user, options.password)
        accounts = get_accounts(web_driver, options.end_date, **options_kwargs)
        logger.info('Processing accounts: %s', ', '.join(accounts.keys()))
        web_driver.execute_script("sendMenuRequest('transactionHistorySelectAccount.ctl');")
        logger.debug('got transactions page')
        closing_balances = {}
        with open('{}-Closing Balances.csv'.format(options.end_date.strftime('%Y%m%d')), 'w') as out_file:
            for key, account in [(k, v) for k, v in accounts.items() if not options.account or (k in options.account)]:
                logger.debug('Doing account:%s', key)
                closing_balances[key] = account.generate_qif(web_driver, options.start_date, options.end_date)
                print>>out_file, '{n}|{a}|{b}'.format(n=account.nick_name, a=account.number if not account.bsb else '{} {}'.format(account.bsb, account.number), b=str(closing_balances[key]))
        logger.info('Closing Balances (as at %s):', options.end_date.strftime('%Y%m%d'))
        for acc, bal in closing_balances.items():
            logger.info('  %s: %s', acc, bal)
    except UnexpectedPageTitle as excp:
        logger.error(excp)
    except Exception as excp:
        traceback.print_exc()
        keep_driver = (keep_driver or options.driver_keep == 'exception')
    finally:
        if keep_driver:
            logger.info('Keeping webdriver instance running')
            time.sleep(6000)
        else:
            web_driver.close()
    logger.info('Finished, %s seconds elapsed', int((datetime.datetime.now() - start_time).total_seconds()))

if __name__ == '__main__':
    main()
