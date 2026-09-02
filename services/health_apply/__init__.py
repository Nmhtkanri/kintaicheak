# -*- coding: utf-8 -*-
"""健康診断申込モード（Google スプレッドシート連携）。

サブモジュールはここから re-export しない。sheets_gateway が google-auth を
関数内で import する設計なので、パッケージを import しただけで Google 系の
依存が要求されないようにしている。
"""
