"""Screening: what never reaches a supplier, and why."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bomlib.prepare import (  # noqa: E402
    DEFAULT_IGNORE_PREFIXES,
    DUPLICATE,
    IGNORED,
    MERGED,
    describe_exclusions,
    matching_prefix,
    normalize_mpn,
    parse_prefixes,
    prepare_lines,
)


def line(row, mpn, quantity=1, **extra):
    entry = {'row': row, 'mpn': mpn, 'quantity': quantity, 'reference': None,
             'manufacturer': None, 'description': None}
    entry.update(extra)
    return entry


class PrefixTests(unittest.TestCase):
    def test_the_four_in_house_prefixes_are_ignored_by_default(self):
        for prefix in ('ASY0', 'CBL0', 'DES0', 'PCB0'):
            result = prepare_lines([line(1, prefix + '-12345')])
            self.assertEqual(result['lines'], [], prefix)
            self.assertEqual(result['excluded'][0]['reason'], IGNORED)

    def test_matching_is_case_insensitive_and_ignores_stray_spacing(self):
        result = prepare_lines([line(1, '  asy0-1 '), line(2, 'Pcb0777')])
        self.assertEqual(result['lines'], [])
        self.assertEqual([e['reason'] for e in result['excluded']], [IGNORED, IGNORED])

    def test_a_prefix_only_counts_at_the_start(self):
        # A real part that merely contains the letters must survive.
        result = prepare_lines([line(1, 'XPCB0-100'), line(2, 'MYASY0')])
        self.assertEqual([r['mpn'] for r in result['lines']], ['XPCB0-100', 'MYASY0'])

    def test_a_digit_matters_because_asy1_is_not_asy0(self):
        result = prepare_lines([line(1, 'ASY1-500')])
        self.assertEqual([r['mpn'] for r in result['lines']], ['ASY1-500'])

    def test_the_reason_names_the_prefix_that_ruled_it_out(self):
        result = prepare_lines([line(4, 'CBL0-9')])
        self.assertIn('CBL0', result['excluded'][0]['detail'])
        self.assertEqual(result['excluded'][0]['row'], 4)

    def test_an_empty_prefix_list_turns_screening_off(self):
        result = prepare_lines([line(1, 'ASY0-1')], ignore_prefixes=[])
        self.assertEqual([r['mpn'] for r in result['lines']], ['ASY0-1'])

    def test_matching_prefix_reports_which_one_matched(self):
        self.assertEqual(matching_prefix('pcb0-1', DEFAULT_IGNORE_PREFIXES), 'PCB0')
        self.assertIsNone(matching_prefix('RC0603', DEFAULT_IGNORE_PREFIXES))
        self.assertIsNone(matching_prefix('', DEFAULT_IGNORE_PREFIXES))


class ParsePrefixTests(unittest.TestCase):
    def test_none_means_unset_and_falls_back_to_the_default(self):
        self.assertEqual(parse_prefixes(None), list(DEFAULT_IGNORE_PREFIXES))

    def test_an_empty_string_is_an_explicit_ignore_nothing(self):
        # Distinct from None: someone who sets the variable to blank means it.
        self.assertEqual(parse_prefixes(''), [])

    def test_commas_and_spaces_both_separate(self):
        self.assertEqual(parse_prefixes('fix0, tst0  jig0'), ['FIX0', 'TST0', 'JIG0'])

    def test_a_list_is_taken_as_given(self):
        self.assertEqual(parse_prefixes(['fix0']), ['FIX0'])


class MergeTests(unittest.TestCase):
    def test_the_same_part_twice_becomes_one_line_with_the_quantities_added(self):
        result = prepare_lines([
            line(1, 'RC0603FR-0710KL', 10, reference='R1'),
            line(5, 'RC0603FR-0710KL', 20, reference='R7'),
        ])
        self.assertEqual(len(result['lines']), 1)
        self.assertEqual(result['lines'][0]['quantity'], 30)
        self.assertEqual(result['lines'][0]['reference'], 'R1, R7')
        self.assertEqual(result['lines'][0]['mergedRows'], [5])
        self.assertEqual(result['excluded'][0]['reason'], MERGED)

    def test_the_merged_line_keeps_a_pointer_back_to_the_row_it_joined(self):
        result = prepare_lines([line(1, 'ABC'), line(9, 'ABC')])
        self.assertIn('row 1', result['excluded'][0]['detail'])

    def test_case_and_spacing_do_not_hide_a_duplicate(self):
        result = prepare_lines([line(1, 'abc 123', 5), line(2, 'ABC  123', 5)])
        self.assertEqual(len(result['lines']), 1)
        self.assertEqual(result['lines'][0]['quantity'], 10)

    def test_reference_designators_are_not_repeated(self):
        result = prepare_lines([
            line(1, 'ABC', 1, reference='R1, R2'),
            line(2, 'ABC', 1, reference='R2, R3'),
        ])
        self.assertEqual(result['lines'][0]['reference'], 'R1, R2, R3')

    def test_a_blank_field_on_the_first_line_is_filled_from_the_second(self):
        result = prepare_lines([
            line(1, 'ABC'),
            line(2, 'ABC', description='10k resistor', manufacturer='Yageo'),
        ])
        self.assertEqual(result['lines'][0]['description'], '10k resistor')
        self.assertEqual(result['lines'][0]['manufacturer'], 'Yageo')

    def test_merging_can_be_turned_off(self):
        result = prepare_lines([line(1, 'ABC', 10), line(2, 'ABC', 20)],
                               merge_duplicates=False)
        self.assertEqual([r['quantity'] for r in result['lines']], [10, 20])
        self.assertEqual(result['excluded'], [])

    def test_the_original_lines_are_not_mutated(self):
        original = line(1, 'ABC', 10)
        prepare_lines([original, line(2, 'ABC', 20)])
        self.assertEqual(original['quantity'], 10)


class ClaimTests(unittest.TestCase):
    def test_a_part_another_bom_already_owns_is_not_looked_up_again(self):
        result = prepare_lines([line(1, 'STM32F103C8T6', 25)],
                               claimed={'STM32F103C8T6': 'Main board'})
        self.assertEqual(result['lines'], [])
        self.assertEqual(result['excluded'][0]['reason'], DUPLICATE)
        self.assertIn('Main board', result['excluded'][0]['detail'])

    def test_claims_are_matched_on_the_normalized_part_number(self):
        result = prepare_lines([line(1, ' stm32f103c8t6 ')],
                               claimed={'STM32F103C8T6': 'Main board'})
        self.assertEqual(result['lines'], [])

    def test_what_survives_is_reported_as_this_boms_claim(self):
        result = prepare_lines([line(1, 'abc'), line(2, 'ASY0-1'), line(3, 'def')])
        self.assertEqual(sorted(result['claimed']), ['ABC', 'DEF'])

    def test_an_in_house_prefix_wins_over_a_claim_so_it_is_never_a_duplicate(self):
        result = prepare_lines([line(1, 'ASY0-1')], claimed={'ASY0-1': 'Main board'})
        self.assertEqual(result['excluded'][0]['reason'], IGNORED)


class DescribeTests(unittest.TestCase):
    def test_nothing_skipped_says_nothing(self):
        self.assertIsNone(describe_exclusions([]))

    def test_each_reason_is_counted_separately(self):
        text = describe_exclusions([
            {'reason': IGNORED}, {'reason': IGNORED},
            {'reason': MERGED}, {'reason': DUPLICATE},
        ])
        self.assertIn('2 in-house', text)
        self.assertIn('1 merged duplicate', text)
        self.assertIn('1 already in another BOM', text)


class NormalizeTests(unittest.TestCase):
    def test_case_and_inner_whitespace_are_folded(self):
        self.assertEqual(normalize_mpn('  grm188  r71h '), 'GRM188 R71H')
        self.assertEqual(normalize_mpn(None), '')


if __name__ == '__main__':
    unittest.main()
