import unittest

from main import OfferToolApp


class BreakEvenCpcTests(unittest.TestCase):
    def make_app(self):
        app = OfferToolApp.__new__(OfferToolApp)
        app.stop_flag = False
        app.log_manage = lambda *args, **kwargs: None
        app.log_debug = lambda *args, **kwargs: None
        app._is_yp_tracking_link = OfferToolApp._is_yp_tracking_link.__get__(app, OfferToolApp)
        app._is_pb_tracking_link = OfferToolApp._is_pb_tracking_link.__get__(app, OfferToolApp)
        app.normalize_sheet_cell_value = OfferToolApp.normalize_sheet_cell_value.__get__(app, OfferToolApp)
        app.extract_country_from_campaign_name = OfferToolApp.extract_country_from_campaign_name.__get__(app, OfferToolApp)
        app.extract_brand_and_country_from_campaign_name = OfferToolApp.extract_brand_and_country_from_campaign_name.__get__(app, OfferToolApp)
        app.normalize_country_code = OfferToolApp.normalize_country_code.__get__(app, OfferToolApp)
        app.normalize_brand_key = OfferToolApp.normalize_brand_key.__get__(app, OfferToolApp)
        app.build_yp_brand_key = OfferToolApp.build_yp_brand_key.__get__(app, OfferToolApp)
        app.build_yp_row_brand_key = OfferToolApp.build_yp_row_brand_key.__get__(app, OfferToolApp)
        app.build_yp_transaction_brand_key = OfferToolApp.build_yp_transaction_brand_key.__get__(app, OfferToolApp)
        app.display_offer_brand_id = OfferToolApp.display_offer_brand_id.__get__(app, OfferToolApp)
        app.campaign_name_matches_brand_country = OfferToolApp.campaign_name_matches_brand_country.__get__(app, OfferToolApp)
        app.parse_commission_value = OfferToolApp.parse_commission_value.__get__(app, OfferToolApp)
        app.extract_tracking_uid_from_link = OfferToolApp.extract_tracking_uid_from_link.__get__(app, OfferToolApp)
        app.build_yp_offer_group_context = OfferToolApp.build_yp_offer_group_context.__get__(app, OfferToolApp)
        app.resolve_yp_campaign_offer_group = OfferToolApp.resolve_yp_campaign_offer_group.__get__(app, OfferToolApp)
        app.summarize_campaign_group_rows = OfferToolApp.summarize_campaign_group_rows.__get__(app, OfferToolApp)
        app.apply_yp_group_commissions_to_campaign_rows = OfferToolApp.apply_yp_group_commissions_to_campaign_rows.__get__(app, OfferToolApp)
        app.build_copied_offer_tracking_link = OfferToolApp.build_copied_offer_tracking_link.__get__(app, OfferToolApp)
        app.is_pb_offer_row = OfferToolApp.is_pb_offer_row.__get__(app, OfferToolApp)
        app.build_copy_offer_match_key = OfferToolApp.build_copy_offer_match_key.__get__(app, OfferToolApp)
        app.copy_offer_has_non_key_data = OfferToolApp.copy_offer_has_non_key_data.__get__(app, OfferToolApp)
        app.format_brand_break_even_cpc = OfferToolApp.format_brand_break_even_cpc.__get__(app, OfferToolApp)
        app.build_ads_brand_campaign_sheet_maps = OfferToolApp.build_ads_brand_campaign_sheet_maps.__get__(app, OfferToolApp)
        app.detect_platform_from_tracking_link = OfferToolApp.detect_platform_from_tracking_link.__get__(app, OfferToolApp)
        app.infer_campaign_platform = OfferToolApp.infer_campaign_platform.__get__(app, OfferToolApp)
        app.build_offer_tracking_link_maps_for_ads_brand = OfferToolApp.build_offer_tracking_link_maps_for_ads_brand.__get__(app, OfferToolApp)
        app.get_ads_brand_platforms_for_campaigns = OfferToolApp.get_ads_brand_platforms_for_campaigns.__get__(app, OfferToolApp)
        app.build_ads_brand_sheet_rows = OfferToolApp.build_ads_brand_sheet_rows.__get__(app, OfferToolApp)
        app.build_offer_group_lookup_for_campaign_summary = OfferToolApp.build_offer_group_lookup_for_campaign_summary.__get__(app, OfferToolApp)
        app.resolve_campaign_summary_brand_name = OfferToolApp.resolve_campaign_summary_brand_name.__get__(app, OfferToolApp)
        app.parse_google_mcc_ids = OfferToolApp.parse_google_mcc_ids.__get__(app, OfferToolApp)
        app.build_break_even_brand_country_offer_index = OfferToolApp.build_break_even_brand_country_offer_index.__get__(app, OfferToolApp)
        app.build_brand_country_lookup_maps = OfferToolApp.build_brand_country_lookup_maps.__get__(app, OfferToolApp)
        app.resolve_yp_brand_country_entry = OfferToolApp.resolve_yp_brand_country_entry.__get__(app, OfferToolApp)
        return app

    def test_break_even_totals_should_combine_pb_and_yp_commission(self):
        app = self.make_app()

        offer_row_commissions = {
            10: {
                'asin': 'ASINPB1',
                'brand_id': '1001',
                'brand_name': 'Anker',
                'country': 'US',
                'is_yp': False,
                'commission': 12.0,
            },
            11: {
                'asin': 'ASINYP1',
                'brand_id': '1001',
                'brand_name': 'Anker',
                'country': 'US',
                'is_yp': True,
                'commission': 8.0,
            },
        }
        offer_group_summary_by_key = {}
        mcc_campaigns = {
            '2037-2589-anker-US-Search-1': {'clicks': 5, '_traffic_source': 'pb'},
            '2037-2589-anker-US-Search-2': {'clicks': 15, '_traffic_source': 'yp'},
        }

        totals = app.build_break_even_brand_country_totals(
            offer_row_commissions=offer_row_commissions,
            offer_group_summary_by_key=offer_group_summary_by_key,
            mcc_campaigns=mcc_campaigns,
        )

        self.assertIn(('anker', 'US'), totals)
        self.assertEqual(20.0, totals[('anker', 'US')]['commission'])
        self.assertEqual(20, totals[('anker', 'US')]['clicks'])

    def test_apply_break_even_cpc_should_ignore_platform_split(self):
        app = self.make_app()

        mcc_campaigns = {
            '2037-2589-anker-US-Search-1': {'clicks': 5, '_traffic_source': 'pb'},
            '2037-2589-anker-US-Search-2': {'clicks': 15, '_traffic_source': 'yp'},
        }
        updates = [
            {'campaign_name': '2037-2589-anker-US-Search-1', '状态': '投放中'},
            {'campaign_name': '2037-2589-anker-US-Search-2', '状态': '广告系列暂停中2026-5-11'},
        ]
        new_rows = []
        brand_country_totals = {
            ('anker', 'US'): {'brand': 'Anker', 'country': 'US', 'commission': 20.0, 'clicks': 20},
        }

        applied = app.apply_brand_break_even_cpc(updates, new_rows, mcc_campaigns, brand_country_totals)

        self.assertEqual(2, applied)
        self.assertEqual('$1.00（$20.00/20）', updates[0]['品牌收支平衡CPC'])
        self.assertEqual('$1.00（$20.00/20）', updates[1]['品牌收支平衡CPC'])

    def test_yp_country_resolution_requires_unique_offer_when_api_has_no_country(self):
        app = self.make_app()
        offer_index = app.build_break_even_brand_country_offer_index([
            {'状态': '', '品牌名称': 'Anker', '品牌ID': '1001', '国家代码': 'US', 'ASIN': 'B0TEST0001'},
            {'状态': '', '品牌名称': 'Anker', '品牌ID': '1001', '国家代码': 'DE', 'ASIN': 'B0TEST0001'},
        ])
        lookup = app.build_brand_country_lookup_maps(offer_index)

        resolved = app.resolve_yp_brand_country_entry({
            'asin': 'B0TEST0001',
            'brand_id': '1001',
            'status': 'approved',
            'sale_comm': 10,
        }, lookup)

        self.assertIsNone(resolved)

    def test_yp_resolution_should_fallback_by_advert_name_asin_country(self):
        app = self.make_app()
        offer_index = app.build_break_even_brand_country_offer_index([
            {'状态': '', '品牌名称': 'Bcan', '品牌ID': '8888', '国家代码': 'US', 'ASIN': 'B0BCAN0001'},
            {'状态': '', '品牌名称': 'Other', '品牌ID': '9999', '国家代码': 'US', 'ASIN': 'B0BCAN0001'},
        ])
        lookup = app.build_brand_country_lookup_maps(offer_index)

        resolved = app.resolve_yp_brand_country_entry({
            'asin': 'B0BCAN0001',
            'advert_name': 'Bcan',
            'customer_country': 'US',
            'sale_comm': 10,
        }, lookup)

        self.assertIsNotNone(resolved)
        self.assertEqual('bcan', resolved['brand_lower'])
        self.assertEqual({'8888'}, resolved['brand_ids'])

    def test_yp_resolution_should_fallback_by_advert_name_asin_without_country(self):
        app = self.make_app()
        offer_index = app.build_break_even_brand_country_offer_index([
            {'状态': '', '品牌名称': 'Merach', '品牌ID': '', '国家代码': 'US', 'ASIN': 'B0F8J55GT8'},
        ])
        lookup = app.build_brand_country_lookup_maps(offer_index)

        resolved = app.resolve_yp_brand_country_entry({
            'id': '2026-05-24T00:00:00.000Z_3d869c2e7faa4abc_B0F8J55GT8',
            'advert_id': '380945',
            'advert_name': 'Merach',
            'sale_comm': 12.34,
        }, lookup)

        self.assertIsNotNone(resolved)
        self.assertEqual('merach', resolved['brand_lower'])
        self.assertEqual('US', resolved['country'])
        self.assertEqual('name:merach', app.build_yp_transaction_brand_key({
            'advert_name': 'Merach',
        }, resolved))

    def test_yp_resolution_should_fallback_by_unique_brand_when_asin_missing_from_sheet(self):
        app = self.make_app()
        offer_index = app.build_break_even_brand_country_offer_index([
            {'状态': '', '品牌名称': 'Merach', '品牌ID': '', '国家代码': 'US', 'ASIN': 'B0F8J55GT8'},
        ])
        lookup = app.build_brand_country_lookup_maps(offer_index)

        resolved = app.resolve_yp_brand_country_entry({
            'id': '2026-05-15T00:00:00.000Z_a70c1370c637d17d_B0DPW7JG4K',
            'advert_id': '380945',
            'advert_name': 'Merach',
            'sale_comm': 15.84,
        }, lookup)

        self.assertIsNotNone(resolved)
        self.assertEqual('merach', resolved['brand_lower'])
        self.assertEqual('US', resolved['country'])

    def test_resolve_yp_campaign_offer_group_can_use_existing_tracking_link(self):
        app = self.make_app()
        feishu_data = [
            {
                '状态': '投放中',
                '品牌名称': 'Anker',
                '品牌ID': '1001',
                '国家代码': 'US',
                'ASIN': 'B0TEST0001',
                '投放链接': 'https://yeahpromos.com/track?pid=abc123&u1={tag1}',
                '广告系列名称': '',
            }
        ]
        offer_context = app.build_yp_offer_group_context(feishu_data, offer_row_commissions={})

        offer_key, summary = app.resolve_yp_campaign_offer_group(
            campaign_name='Old YP Campaign',
            campaign_info={'campaign_id': '', 'asin': '', 'country': '', 'tracking_link': 'https://yeahpromos.com/track?pid=abc123&u1={tag1}'},
            offer_context=offer_context,
            tracking_link='https://yeahpromos.com/track?pid=abc123&u1={tag1}',
        )

        self.assertEqual(('B0TEST0001', '1001', 'US'), offer_key)
        self.assertTrue(summary.get('is_yp'))

    def test_summarize_campaign_group_rows_should_aggregate_display_metrics(self):
        app = self.make_app()
        row_by_index = {
            10: {
                '广告系列总花费': '$12.50',
                '总佣金': '$20.00',
                '总点击数': '5',
                '佣金ASIN': 'B0TEST0001_US',
                '品牌收支平衡CPC': '$1.50',
                '新增广告系列花费': '$2.50',
                '新增佣金': '$4.00',
            },
            11: {
                '广告系列总花费': '$7.50',
                '总佣金': '$8.00',
                '总点击数': '3',
                '佣金ASIN': 'B0TEST0002_US, B0TEST0001_US',
                '品牌收支平衡CPC': '$1.50',
                '新增广告系列花费': '$1.00',
                '新增佣金': '$2.00',
            },
        }

        metrics = app.summarize_campaign_group_rows(row_by_index, [10, 11])

        self.assertEqual(20.0, metrics['total_cost'])
        self.assertEqual(28.0, metrics['total_commission'])
        self.assertEqual(8, metrics['total_clicks'])
        self.assertEqual(['B0TEST0001_US', 'B0TEST0002_US'], metrics['commission_asins'])
        self.assertEqual('$1.50', metrics['brand_break_even_cpc'])
        self.assertEqual(3.5, metrics['increment_cost'])
        self.assertEqual(6.0, metrics['increment_commission'])

    def test_apply_yp_group_commissions_to_campaign_rows_should_fill_single_campaign_group(self):
        app = self.make_app()
        updates = [{
            'campaign_name': 'Sihoo_US_6773_11257_20260507161819286',
            '总佣金': '$0.00',
            '广告系列总花费': '$10.00',
            'ROI': '0.0',
            '佣金ASIN': '',
            '_offer_group_key': ('B0TEST0001', '1001', 'US'),
        }]
        new_rows = []
        result = app.apply_yp_group_commissions_to_campaign_rows(
            updates=updates,
            new_rows=new_rows,
            offer_group_summary_by_key={
                ('B0TEST0001', '1001', 'US'): {
                    'commission': 66.94,
                    'is_yp': True,
                }
            },
            existing_campaign_commission={'Sihoo_US_6773_11257_20260507161819286': 0.0},
            existing_campaign_cost={'Sihoo_US_6773_11257_20260507161819286': 10.0},
        )

        self.assertEqual(1, result['applied_rows'])
        self.assertEqual('$66.94', updates[0]['总佣金'])
        self.assertEqual('$66.94', updates[0]['新增佣金'])
        self.assertEqual('6.7', updates[0]['ROI'])
        self.assertEqual('B0TEST0001_US', updates[0]['佣金ASIN'])

    def test_build_copied_offer_tracking_link_should_keep_existing_yp_link(self):
        app = self.make_app()
        source_link = 'https://yeahpromos.com/track?pid=abc123&u1={tag1}'
        app.generate_random_uid = lambda: self.fail('YP link should not request a new uid')
        app.get_partnerboost_link = lambda asin, country, uid: self.fail('YP link should not call PB link API')

        new_link, generated_new_link, uid = app.build_copied_offer_tracking_link('B0TEST0001', 'US', source_link)

        self.assertEqual(source_link, new_link)
        self.assertFalse(generated_new_link)
        self.assertEqual('', uid)

    def test_build_copied_offer_tracking_link_should_generate_new_pb_link(self):
        app = self.make_app()
        app.generate_random_uid = lambda: 'abc1234'
        app.get_partnerboost_link = lambda asin, country, uid: f'https://pboost.me/link?asin={asin}&country={country}&uid={uid}'

        new_link, generated_new_link, uid = app.build_copied_offer_tracking_link(
            'B0TEST0001', 'US', 'https://pboost.me/original-link'
        )

        self.assertTrue(generated_new_link)
        self.assertEqual('abc1234', uid)
        self.assertEqual('https://pboost.me/link?asin=B0TEST0001&country=US&uid=abc1234', new_link)

    def test_is_pb_offer_row_should_skip_non_pb_rows(self):
        app = self.make_app()
        col_indices = {'投放链接': 0, '状态': 1}

        self.assertFalse(app.is_pb_offer_row(['https://yeahpromos.com/track?pid=abc123&u1={tag1}', ''], col_indices))
        self.assertFalse(app.is_pb_offer_row(['https://pboost.me/abc123', '相同offer统计行'], col_indices))
        self.assertTrue(app.is_pb_offer_row(['https://pboost.me/abc123', ''], col_indices))
        self.assertFalse(app.is_pb_offer_row(['', ''], col_indices))

    def test_is_pb_offer_row_should_require_explicit_pb_link(self):
        app = self.make_app()
        col_indices = {'投放链接': 0, '状态': 1}

        self.assertFalse(app.is_pb_offer_row(['https://example.com/offer', ''], col_indices))
        self.assertTrue(app.is_pb_offer_row(['https://www.pboost.me/abc123', ''], col_indices))

    def test_offer_total_cost_should_write_zero_when_campaign_cost_is_zero(self):
        app = self.make_app()
        campaigns = [{
            'cost_usd': 0.0,
            'status': 'ENABLED',
            'account_id': '123-456-7890',
            'campaign_name': 'TestCampaign',
            'campaign_id': '1',
            'asin': 'B0TEST0001',
            'country': 'US',
        }]
        row = {
            'row_index': 10,
            'ASIN': 'B0TEST0001',
            '国家代码': 'US',
            '状态': '投放中',
            '广告系列总花费': '$12.34',
            '投放链接': 'https://pboost.me/x?uid=abc1234',
            '品牌ID': '1001',
            '品牌名称': 'Brand',
        }

        updates, _ = app.calculate_updates(
            feishu_data=[row],
            asin_country_campaigns={('B0TEST0001', 'US'): campaigns},
            asin_country_commission={},
            row_campaigns={10: campaigns}
        )

        self.assertEqual(1, len(updates))
        self.assertEqual(0.0, updates[0]['total_cost'])

    def test_build_copy_offer_match_key_should_include_brand_id(self):
        app = self.make_app()
        col_indices = {'ASIN': 0, '国家代码': 1, '品牌ID': 2}

        key = app.build_copy_offer_match_key(['B0TEST0001', 'us', '1001'], col_indices)

        self.assertEqual(('B0TEST0001', 'US', '1001'), key)

    def test_build_copy_offer_match_key_should_require_brand_id(self):
        app = self.make_app()
        col_indices = {'ASIN': 0, '国家代码': 1, '品牌ID': 2}

        key = app.build_copy_offer_match_key(['B0TEST0001', 'US', ''], col_indices)

        self.assertIsNone(key)

    def test_copy_offer_has_non_key_data_should_ignore_brand_id(self):
        app = self.make_app()
        col_indices = {'品牌ID': 0, '品牌名称': 1, '佣金': 2}
        copy_fields = ['品牌名称', '品牌ID', '佣金']

        result = app.copy_offer_has_non_key_data(['1001', '', ''], col_indices, copy_fields)

        self.assertFalse(result)

    def test_copy_offer_has_non_key_data_should_detect_real_offer_fields(self):
        app = self.make_app()
        col_indices = {'品牌ID': 0, '品牌名称': 1, '佣金': 2}
        copy_fields = ['品牌名称', '品牌ID', '佣金']

        result = app.copy_offer_has_non_key_data(['1001', 'Anker', ''], col_indices, copy_fields)

        self.assertTrue(result)

    def test_format_brand_break_even_cpc_should_include_commission_and_clicks(self):
        app = self.make_app()

        result = app.format_brand_break_even_cpc(100, 50)

        self.assertEqual('$2.00（$100.00/50）', result)

    def test_format_brand_break_even_cpc_should_handle_zero_clicks(self):
        app = self.make_app()

        result = app.format_brand_break_even_cpc(100, 0)

        self.assertEqual('$0.00（$100.00/0）', result)

    def test_normalize_brand_key_should_strip_country_suffix_or_prefix(self):
        app = self.make_app()

        self.assertEqual('aosu', app.normalize_brand_key('Aosu CA'))
        self.assertEqual('eureka', app.normalize_brand_key('Eureka France'))
        self.assertEqual('anker', app.normalize_brand_key('DE-Anker'))
        self.assertEqual('maono', app.normalize_brand_key('Maono_GB'))
        self.assertEqual('ogery', app.normalize_brand_key('OGERY-Outdoors'))

    def test_extract_country_from_campaign_name_should_normalize_gb_to_uk(self):
        app = self.make_app()

        result = app.extract_country_from_campaign_name('Maono_GB_6124_10248_20260429112122577')

        self.assertEqual('UK', result)

    def test_campaign_name_matches_brand_country_should_tolerate_brand_country_suffixes(self):
        app = self.make_app()

        self.assertTrue(app.campaign_name_matches_brand_country(
            'Eureka_FR_6092_10228_20260428171255765',
            'Eureka France',
            'FR'
        ))
        self.assertTrue(app.campaign_name_matches_brand_country(
            'Maono_GB_6124_10248_20260429112122577',
            'Maono UK',
            'UK'
        ))

    def test_ads_brand_rows_should_include_recent_5_day_total_cost(self):
        app = self.make_app()
        app.calculate_break_even_brand_country_commissions = lambda **kwargs: ({}, {})

        rows = app.build_ads_brand_sheet_rows(
            campaigns=[
                {
                    'account_id': '123',
                    'account_name': 'ads-a',
                    'mcc_id': '2160853519',
                    'mcc_name': 'mcc11',
                    'campaign_name': 'Anker_US_1',
                    'status': 'ENABLED',
                    'brand': 'Anker',
                    'brand_key': 'anker',
                    'country': 'US',
                    'cost_usd': 100.0,
                    'recent_5_day_cost_usd': 12.34,
                    'clicks': 10,
                    'preset_cpc_usd': 1.0,
                    'avg_cpc_usd': 10.0,
                },
                {
                    'account_id': '123',
                    'account_name': 'ads-a',
                    'mcc_id': '2160853519',
                    'mcc_name': 'mcc11',
                    'campaign_name': 'Anker_US_2',
                    'status': 'PAUSED',
                    'brand': 'Anker',
                    'brand_key': 'anker',
                    'country': 'US',
                    'cost_usd': 20.0,
                    'recent_5_day_cost_usd': 5.0,
                    'clicks': 2,
                    'preset_cpc_usd': 2.0,
                    'avg_cpc_usd': 10.0,
                },
            ],
            campaign_sheet_rows=[],
            feishu_data=[],
            commission_data=[],
            yp_commission_data=[],
            start_date_str='2026-05-01',
            end_date_str='2026-05-21'
        )

        self.assertEqual(1, len(rows))
        self.assertEqual('$17.34', rows[0]['近5日总花费'])

    def test_ads_brand_platform_should_not_default_to_pb_without_link(self):
        app = self.make_app()
        app.calculate_break_even_brand_country_commissions = lambda **kwargs: ({}, {})

        rows = app.build_ads_brand_sheet_rows(
            campaigns=[{
                'account_id': '123',
                'account_name': 'ads-a',
                'mcc_id': '2160853519',
                'mcc_name': 'mcc11',
                'campaign_name': 'Anker_US_1',
                'status': 'ENABLED',
                'brand': 'Anker',
                'brand_key': 'anker',
                'country': 'US',
                'cost_usd': 100.0,
                'recent_5_day_cost_usd': 0.0,
            }],
            campaign_sheet_rows=[],
            feishu_data=[],
            commission_data=[],
            yp_commission_data=[],
            start_date_str='2026-05-01',
            end_date_str='2026-05-21'
        )

        self.assertEqual('', rows[0]['平台'])

    def test_ads_brand_platform_should_use_offer_tracking_link_first(self):
        app = self.make_app()
        app.calculate_break_even_brand_country_commissions = lambda **kwargs: ({}, {})

        rows = app.build_ads_brand_sheet_rows(
            campaigns=[{
                'account_id': '123',
                'account_name': 'ads-a',
                'mcc_id': '2160853519',
                'mcc_name': 'mcc11',
                'campaign_name': 'Anker_US_1',
                'status': 'ENABLED',
                'brand': 'Anker',
                'brand_key': 'anker',
                'country': 'US',
                'cost_usd': 100.0,
                'recent_5_day_cost_usd': 0.0,
            }],
            campaign_sheet_rows=[{
                '广告系列名称': 'Anker_US_1',
                '投放链接': '',
            }],
            feishu_data=[{
                '状态': '投放中',
                '品牌名称': 'Anker',
                '国家代码': 'US',
                '广告系列名称': 'Anker_US_1',
                '投放链接': 'https://yeahpromos.com/track?pid=abc123&u1={tag1}',
            }],
            commission_data=[],
            yp_commission_data=[],
            start_date_str='2026-05-01',
            end_date_str='2026-05-21'
        )

        self.assertEqual('yp', rows[0]['平台'])

    def test_ads_brand_rows_should_include_empty_account_without_active_or_paused_campaigns(self):
        app = self.make_app()
        app.calculate_break_even_brand_country_commissions = lambda **kwargs: ({}, {})

        rows = app.build_ads_brand_sheet_rows(
            campaigns=[
                {
                    'account_id': '456',
                    'account_name': 'ads-empty',
                    'mcc_id': '2160853519',
                    'mcc_name': 'mcc11',
                    'campaign_name': 'Anker_US_removed',
                    'status': 'REMOVED',
                    'brand': 'Anker',
                    'brand_key': 'anker',
                    'country': 'US',
                    'cost_usd': 10.0,
                    'recent_5_day_cost_usd': 1.0,
                },
            ],
            campaign_sheet_rows=[],
            feishu_data=[],
            commission_data=[],
            yp_commission_data=[],
            start_date_str='2026-05-01',
            end_date_str='2026-05-21',
            accounts=[
                {
                    'account_id': '456',
                    'account_name': 'ads-empty',
                    'mcc_id': '2160853519',
                    'mcc_name': 'mcc11',
                }
            ]
        )

        self.assertEqual(1, len(rows))
        self.assertEqual('ads-empty', rows[0]['ads'])
        self.assertEqual('456', rows[0]['ID'])
        self.assertEqual('mcc11', rows[0]['所属mcc'])
        self.assertEqual('2160853519', rows[0]['mccID'])
        self.assertNotIn('在投品牌名', rows[0])
        self.assertNotIn('广告系列总花费', rows[0])

    def test_campaign_summary_brand_name_should_fallback_by_asin_country(self):
        app = self.make_app()
        lookup = app.build_offer_group_lookup_for_campaign_summary([
            {
                '状态': '投放中',
                'ASIN': 'B0C3T865C2',
                '品牌ID': '363072',
                '品牌名称': 'SIHOO US',
                '国家代码': 'US',
                '产品链接': 'https://www.amazon.com/dp/B0C3T865C2',
                '投放链接': 'https://pboost.me/demo',
            }
        ])

        brand_name = app.resolve_campaign_summary_brand_name(
            offer_key=('B0C3T865C2', '126677', 'US'),
            offer_summary={},
            yp_offer_context={},
            offer_group_lookup=lookup,
        )

        self.assertEqual('SIHOO US', brand_name)

    def test_parse_google_mcc_ids_should_accept_list_commas_and_newlines(self):
        app = self.make_app()

        self.assertEqual(
            ['2160853519', '6885177935', '1234567890'],
            app.parse_google_mcc_ids('216-085-3519, 6885177935\n1234567890;2160853519')
        )

    def test_grouped_increment_commission_uses_campaign_group_key(self):
        app = self.make_app()
        rows = [{
            'campaign_name': 'Reolink_US_8388_13291_20260521153755612',
            '总佣金': '$10.00',
            '_offer_group_key': ('B0F9F4D9XS', '136519', 'US', 'adg:1'),
        }]

        app.apply_campaign_increment_commission_delta(
            rows,
            {'campaign_commission': {'Reolink_US_8388_13291_20260521153755612': 2.0}},
            grouped_offer_keys={('B0F9F4D9XS', '136519', 'US')}
        )

        self.assertEqual('↑', rows[0]['总佣金'])
        self.assertEqual('↑', rows[0]['新增佣金'])

    def test_increment_delta_should_overwrite_explicit_increment_commission(self):
        app = self.make_app()
        rows = [{
            'campaign_name': 'Merach_US_1',
            '总佣金': '$56.64',
            '新增佣金': '$15.84',
        }]

        app.apply_campaign_increment_commission_delta(
            rows,
            {'campaign_commission': {'Merach_US_1': 56.64}},
            grouped_offer_keys=set()
        )

        self.assertEqual('$0.00', rows[0]['新增佣金'])

    def test_campaign_background_styles_are_added_for_new_rows(self):
        app = self.make_app()
        captured = {}
        app.log_manage = lambda *args, **kwargs: None
        app._ensure_sheet_rows = lambda *args, **kwargs: None
        app.apply_campaigns_style_updates = lambda token, spreadsheet_token, sheet_id, styles: captured.setdefault('styles', styles)

        class Response:
            def json(self):
                return {'code': 0}

        original_post = __import__('requests').post
        try:
            __import__('requests').post = lambda *args, **kwargs: Response()
            app.apply_campaigns_sheet_updates(
                token='token',
                spreadsheet_token='sheet',
                sheet_id='XrkOF7',
                updates=[],
                new_rows=[{
                    '状态': '投放中',
                    '广告系列名称': 'NewCampaign',
                    '预设CPC': '$1.00',
                    'CPC': '$0.50',
                    '品牌收支平衡CPC': '$2.00',
                    '广告系列总花费': '$3.00',
                    '总佣金': '$4.00',
                    'status_color': 'green',
                }],
                column_map={
                    '状态': 'A',
                    '广告系列名称': 'B',
                    '预设CPC': 'I',
                    'CPC': 'J',
                    '品牌收支平衡CPC': 'K',
                    '广告系列总花费': 'O',
                    '总佣金': 'P',
                },
                first_empty_row=10
            )
        finally:
            __import__('requests').post = original_post

        bg_by_column = {
            item['column']: item['background_color']
            for item in captured.get('styles', [])
            if item.get('background_color')
        }
        self.assertEqual('#F8F9FA', bg_by_column['I'])
        self.assertEqual('#FAF1D1', bg_by_column['J'])
        self.assertEqual('#FFF258', bg_by_column['K'])
        self.assertEqual('#FBBFBC', bg_by_column['O'])
        self.assertEqual('#D9F5D6', bg_by_column['P'])

    def test_campaign_metric_background_styles_include_summary_rows(self):
        app = self.make_app()

        styles = app.build_campaign_metric_background_style_updates(
            rows=[
                {'row_index': 2, '状态': '相同offer统计行', '广告系列名称': 'Brand | ASIN | 1 | US | pb | 1 | 0 | 0 | -'},
                {'row_index': 3, '状态': '投放中', '广告系列名称': 'Brand_US_1'},
            ],
            column_map={
                '预设CPC': 'I',
                'CPC': 'J',
                '品牌收支平衡CPC': 'K',
                '广告系列总花费': 'O',
                '总佣金': 'P',
            }
        )

        self.assertEqual({2, 3}, {item['row_index'] for item in styles})
        self.assertEqual(10, len(styles))


if __name__ == '__main__':
    unittest.main()
