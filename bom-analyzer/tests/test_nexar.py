"""Nexar client: token handling, and reading a GraphQL answer.

The GraphQL query itself could not be checked against a live schema — the
sandbox this was built in has no route to api.nexar.com — so these tests pin
the parts that are ours: that a token is minted once and reused, that a
GraphQL errors array is raised rather than read as "no alternatives", and that
a response is flattened defensively. If the schema turns out to differ, the
recorded response below is the one thing to correct.
"""

import json
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib import nexar  # noqa: E402
from bomlib.http_client import HttpError  # noqa: E402


def part_node(mpn, manufacturer='Acme', similar=None, **extra):
    node = {
        'mpn': mpn,
        'manufacturer': {'name': manufacturer},
        'shortDescription': 'A part that does a thing',
        'octopartUrl': 'https://octopart.com/' + mpn,
        'totalAvail': 4200,
        'estimatedFactoryLeadDays': 84,
        'medianPrice1000': {'price': 0.42, 'currency': 'USD'},
        'bestDatasheet': {'url': 'https://example.com/%s.pdf' % mpn},
        'specs': [
            {'attribute': {'name': 'Number of Channels', 'shortname': 'channels'},
             'displayValue': '2'},
            {'attribute': {'name': 'Supply Voltage'}, 'displayValue': '3.3 V'},
        ],
    }
    if similar is not None:
        node['similarParts'] = similar
    node.update(extra)
    return node


def response(*parts):
    return {'data': {'supSearchMpn': {'results': [{'part': p} for p in parts]}}}


class FakeTransport:
    """Stands in for request_json, recording what the client asked for."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = []

    def __call__(self, url, method='GET', headers=None, body=None, **kwargs):
        self.calls.append({'url': url, 'method': method, 'headers': headers or {}, 'body': body})
        answer = self.answers.pop(0) if self.answers else {}
        if callable(answer):
            return answer()
        return {'status': 200, 'data': answer}


class ClientTests(unittest.TestCase):
    def setUp(self):
        self._original = nexar.request_json
        self.token = {'access_token': 'tok-1', 'expires_in': 3600}

    def tearDown(self):
        nexar.request_json = self._original

    def client(self, *answers, **kwargs):
        transport = FakeTransport(self.token, *answers)
        nexar.request_json = transport
        client = nexar.NexarClient(client_id='id', client_secret='secret', **kwargs)
        return client, transport

    def test_credentials_are_what_makes_it_configured(self):
        self.assertFalse(nexar.NexarClient().configured)
        self.assertFalse(nexar.NexarClient(client_id='id').configured)
        self.assertTrue(nexar.NexarClient(client_id='id', client_secret='s').configured)

    def test_the_token_request_is_client_credentials_with_a_scope(self):
        client, transport = self.client(response())
        client.find_alternatives({'mpn': 'ABC123'})
        body = transport.calls[0]['body']
        self.assertIn('grant_type=client_credentials', body)
        self.assertIn('client_id=id', body)
        self.assertIn('scope=supply.domain', body)
        self.assertEqual(transport.calls[0]['url'], nexar.TOKEN_URL)

    def test_the_token_is_minted_once_and_reused(self):
        client, transport = self.client(response(), response())
        client.find_alternatives({'mpn': 'AAA'})
        client.find_alternatives({'mpn': 'BBB'})
        token_calls = [c for c in transport.calls if c['url'] == nexar.TOKEN_URL]
        self.assertEqual(len(token_calls), 1)

    def test_a_token_response_without_a_token_is_an_error_not_a_blank_query(self):
        nexar.request_json = FakeTransport({'error': 'invalid_client'})
        client = nexar.NexarClient(client_id='id', client_secret='bad')
        with self.assertRaises(HttpError) as caught:
            client.find_alternatives({'mpn': 'ABC123'})
        self.assertIn('access_token', str(caught.exception))

    def test_the_query_is_posted_as_graphql_with_the_bearer_token(self):
        client, transport = self.client(response())
        client.find_alternatives({'mpn': 'ABC123', 'manufacturer': 'Acme'})
        call = transport.calls[1]
        self.assertEqual(call['url'], nexar.API_URL)
        self.assertEqual(call['method'], 'POST')
        self.assertEqual(call['headers']['Authorization'], 'Bearer tok-1')
        payload = json.loads(call['body'])
        self.assertIn('supSearchMpn', payload['query'])
        # The manufacturer narrows a part number several makers use.
        self.assertEqual(payload['variables']['q'], 'ABC123 Acme')

    def test_a_graphql_errors_array_is_raised_rather_than_read_as_no_alternatives(self):
        # GraphQL answers 200 with errors, so a schema mismatch would otherwise
        # look exactly like a part with nothing to replace it.
        client, _ = self.client({'errors': [{'message': "Cannot query field 'similarParts'"}]})
        with self.assertRaises(HttpError) as caught:
            client.find_alternatives({'mpn': 'ABC123'})
        self.assertIn('similarParts', str(caught.exception))
        self.assertIn('rejected the query', str(caught.exception))

    def test_alternatives_come_back_flattened(self):
        client, _ = self.client(response(part_node(
            'ABC123', similar=[part_node('DEF456', 'Beta Corp'), part_node('GHI789')])))
        found = client.find_alternatives({'mpn': 'ABC123'})

        self.assertEqual(found['matched']['mpn'], 'ABC123')
        self.assertEqual([a['mpn'] for a in found['alternatives']], ['DEF456', 'GHI789'])
        first = found['alternatives'][0]
        self.assertEqual(first['manufacturer'], 'Beta Corp')
        self.assertEqual(first['stock'], 4200)
        self.assertEqual(first['medianPrice'], 0.42)
        self.assertEqual(first['currency'], 'USD')
        self.assertEqual(first['leadDays'], 84)
        self.assertEqual(first['datasheetUrl'], 'https://example.com/DEF456.pdf')
        # The shortname is preferred, with the long name as the fallback.
        self.assertEqual([s['name'] for s in first['specs']], ['channels', 'Supply Voltage'])

    def test_a_part_with_no_match_is_not_an_error(self):
        client, _ = self.client(response())
        found = client.find_alternatives({'mpn': 'NOSUCHPART'})
        self.assertIsNone(found['matched'])
        self.assertEqual(found['alternatives'], [])

    def test_a_match_with_nothing_similar_is_not_an_error_either(self):
        client, _ = self.client(response(part_node('ABC123', similar=[])))
        found = client.find_alternatives({'mpn': 'ABC123'})
        self.assertEqual(found['matched']['mpn'], 'ABC123')
        self.assertEqual(found['alternatives'], [])

    def test_the_part_itself_is_never_offered_as_its_own_alternative(self):
        client, _ = self.client(response(part_node(
            'ABC123', similar=[part_node('abc123'), part_node('DEF456')])))
        found = client.find_alternatives({'mpn': 'ABC123'})
        self.assertEqual([a['mpn'] for a in found['alternatives']], ['DEF456'])

    def test_a_duplicate_suggestion_is_listed_once(self):
        client, _ = self.client(response(part_node(
            'ABC123', similar=[part_node('DEF456'), part_node('DEF456')])))
        found = client.find_alternatives({'mpn': 'ABC123'})
        self.assertEqual([a['mpn'] for a in found['alternatives']], ['DEF456'])

    def test_the_number_of_alternatives_is_capped(self):
        similar = [part_node('ALT%03d' % i) for i in range(30)]
        client, _ = self.client(response(part_node('ABC123', similar=similar)),
                                alternatives_limit=5)
        found = client.find_alternatives({'mpn': 'ABC123'})
        self.assertEqual(len(found['alternatives']), 5)

    def test_missing_optional_relations_are_read_as_missing_not_as_a_failure(self):
        bare = {'mpn': 'BARE1', 'similarParts': [{'mpn': 'BARE2'}]}
        client, _ = self.client(response(bare))
        found = client.find_alternatives({'mpn': 'BARE1'})
        alternative = found['alternatives'][0]
        self.assertEqual(alternative['mpn'], 'BARE2')
        for field in ('manufacturer', 'description', 'url', 'datasheetUrl',
                      'stock', 'leadDays', 'medianPrice', 'currency'):
            self.assertIsNone(alternative[field], field)
        self.assertEqual(alternative['specs'], [])

    def test_a_node_with_no_part_number_is_not_a_part(self):
        self.assertIsNone(nexar.part_from_node({'mpn': '   '}))
        self.assertIsNone(nexar.part_from_node(None))
        self.assertIsNone(nexar.part_from_node('not a dict'))

    def test_an_empty_part_number_asks_nothing(self):
        client, transport = self.client()
        found = client.find_alternatives({'mpn': '  '})
        self.assertEqual(found, {'matched': None, 'alternatives': []})
        self.assertEqual(transport.calls, [])


class TokenFailureTests(unittest.TestCase):
    """A 400 from the token endpoint has to say which 400 it was."""

    def setUp(self):
        self._original = nexar.request_json

    def tearDown(self):
        nexar.request_json = self._original

    def refuse(self, body, status=400):
        def transport(url, **kwargs):
            raise HttpError('HTTP %d from identity.nexar.com' % status, status, body)
        nexar.request_json = transport

    def test_the_reason_reaches_the_message_not_just_the_status(self):
        # "HTTP 400 from identity.nexar.com" is the one thing nobody can act on.
        self.refuse({'error': 'invalid_client'})
        client = nexar.NexarClient(client_id='id', client_secret='wrong')
        with self.assertRaises(HttpError) as caught:
            client.get_token()
        message = str(caught.exception)
        self.assertIn('invalid_client', message)
        self.assertIn('NEXAR_CLIENT_SECRET', message)
        self.assertNotEqual(message, 'HTTP 400 from identity.nexar.com')

    def test_the_endpoints_own_description_is_kept(self):
        self.refuse({'error': 'invalid_client', 'error_description': 'Client is disabled'})
        client = nexar.NexarClient(client_id='id', client_secret='s')
        with self.assertRaises(HttpError) as caught:
            client.get_token()
        self.assertIn('Client is disabled', str(caught.exception))

    def test_the_scope_asked_for_is_named(self):
        self.refuse({'error': 'invalid_scope'})
        client = nexar.NexarClient(client_id='id', client_secret='s', scope='design.domain')
        with self.assertRaises(HttpError) as caught:
            client.get_token()
        self.assertIn('design.domain', str(caught.exception))

    def test_an_unparseable_body_still_says_what_came_back(self):
        self.refuse('Bad Request')
        client = nexar.NexarClient(client_id='id', client_secret='s')
        with self.assertRaises(HttpError) as caught:
            client.get_token()
        self.assertIn('Bad Request', str(caught.exception))

    def test_a_refused_scope_is_retried_without_one(self):
        # The commonest cause: the application was never granted the scope.
        # Nexar issues a usable token without one, so the run should continue.
        asked = []

        def transport(url, method='GET', headers=None, body=None, **kwargs):
            asked.append(body)
            if 'scope=' in body:
                raise HttpError('HTTP 400', 400, {'error': 'invalid_scope'})
            return {'status': 200, 'data': {'access_token': 'tok', 'expires_in': 3600}}

        nexar.request_json = transport
        client = nexar.NexarClient(client_id='id', client_secret='s')
        self.assertEqual(client.get_token(), 'tok')
        self.assertEqual(len(asked), 2)
        self.assertIn('scope=supply.domain', asked[0])
        self.assertNotIn('scope=', asked[1])
        # And it records which one worked, so the app can say so.
        self.assertIsNone(client.scope_used)

    def test_any_other_refusal_is_not_retried(self):
        # Retrying bad credentials without a scope would only fail again, more
        # slowly and with a less useful message.
        asked = []

        def transport(url, method='GET', headers=None, body=None, **kwargs):
            asked.append(body)
            raise HttpError('HTTP 400', 400, {'error': 'invalid_client'})

        nexar.request_json = transport
        client = nexar.NexarClient(client_id='id', client_secret='s')
        with self.assertRaises(HttpError):
            client.get_token()
        self.assertEqual(len(asked), 1)

    def test_no_scope_is_asked_for_when_there_is_none(self):
        sent = []

        def transport(url, method='GET', headers=None, body=None, **kwargs):
            sent.append(body)
            return {'status': 200, 'data': {'access_token': 'tok', 'expires_in': 60}}

        nexar.request_json = transport
        client = nexar.NexarClient(client_id='id', client_secret='s', scope='')
        client.get_token()
        self.assertNotIn('scope', sent[0])


class ScopeConfigTests(unittest.TestCase):
    def test_an_unset_scope_uses_the_default(self):
        self.assertEqual(nexar.client_from_env({}).scope, nexar.SCOPE)

    def test_an_empty_scope_means_ask_for_none(self):
        # Distinct from unset, exactly as the ignore-prefix list is.
        self.assertIsNone(nexar.client_from_env({'NEXAR_SCOPE': ''}).scope)

    def test_a_named_scope_is_used_as_given(self):
        self.assertEqual(nexar.client_from_env({'NEXAR_SCOPE': ' user.access '}).scope,
                         'user.access')


class EnvTests(unittest.TestCase):
    def test_the_client_is_built_from_the_documented_variables(self):
        client = nexar.client_from_env({
            'NEXAR_CLIENT_ID': 'id', 'NEXAR_CLIENT_SECRET': 'secret',
            'NEXAR_ALTERNATIVES_LIMIT': '4',
        })
        self.assertTrue(client.configured)
        self.assertEqual(client.alternatives_limit, 4)
        self.assertEqual(client.scope, nexar.SCOPE)

    def test_the_query_can_be_replaced_without_a_code_change(self):
        # GraphQL schemas move; a wrong field name should not need a release.
        import tempfile
        handle, path = tempfile.mkstemp(suffix='.graphql')
        os.close(handle)
        try:
            with open(path, 'w', encoding='utf-8') as out:
                out.write('query Custom($q: String!, $limit: Int!) { supSearchMpn(q: $q) { x } }')
            client = nexar.client_from_env({
                'NEXAR_CLIENT_ID': 'id', 'NEXAR_CLIENT_SECRET': 's',
                'NEXAR_QUERY_FILE': path,
            })
            self.assertIn('query Custom', client.query)
        finally:
            os.unlink(path)

    def test_no_override_leaves_the_built_in_query(self):
        client = nexar.client_from_env({'NEXAR_CLIENT_ID': 'id', 'NEXAR_CLIENT_SECRET': 's'})
        self.assertIn('supSearchMpn', client.query)


if __name__ == '__main__':
    unittest.main()
