import datetime as dt
import random as rnd
from sys import float_info as sflt
import os
from tqdm import tqdm
from datetime import datetime

import numpy as np
import pandas as pd
# pd.set_option("max_rows", 5)
pd.options.display.max_rows = 20
# pd.set_option("max_columns", 20)
pd.options.display.max_columns = 5

import mplfinance as mpf
import pandas_ta as ta
import vectorbt as vbt

import sys
sys.path.append('examples')
from watchlist import colors, Watchlist # Is this failing? If so, copy it locally. See above.

    
# Function to format Millions
def format_millions(x, pos):
    "The two args are the value and tick position"
    return "%1.1fM" % (x * 1e-6)

# def ctitle(indicator_name, ticker="SPY", length=100):
#     return f"{ticker}: {indicator_name} from {recent_s rtdate} to {recent_enddate} ({length})"

# # All Data: 0, Last Four Years: 0.25, Last Two Years: 0.5, This Year: 1, Last Half Year: 2, Last Quarter: 3
# yearly_divisor = 1
# recent = int(ta.RATE["TRADING_DAYS_PER_YEAR"] / yearly_divisor) if yearly_divisor > 0 else df.shape[0]
# print(recent)
def recent_bars(df, tf: str = "1y"):
    # All Data: 0, Last Four Years: 0.25, Last Two Years: 0.5, This Year: 1, Last Half Year: 2, Last Quarter: 4
    yearly_divisor = {"all": 0, "10y": 0.1, "5y": 0.2, "4y": 0.25, "3y": 1./3, "2y": 0.5, "1y": 1, "6mo": 2, "3mo": 4}
    yd = yearly_divisor[tf] if tf in yearly_divisor.keys() else 0
    return int(ta.RATE["TRADING_DAYS_PER_YEAR"] / yd) if yd > 0 else df.shape[0]

def ta_ylim(series: pd.Series, percent: float = 0.1):
    smin, smax = series.min(), series.max()
    if isinstance(percent, float) and 0 <= float(percent) <= 1:
        y_min = (1 + percent) * smin if smin < 0 else (1 - percent) * smin
        y_max = (1 - percent) * smax if smax < 0 else (1 + percent) * smax
        return (y_min, y_max)
    return (smin, smax)

class Chart(object):
    def __init__(self, df: pd.DataFrame = None, strategy: ta.Strategy = ta.CommonStrategy, *args, **kwargs):
        self.verbose = kwargs.pop("verbose", False)

        if isinstance(df, pd.DataFrame) and df.ta.datetime_ordered:
            self.df = df
            if self.df.name is not None and self.df.name != "":
                df_name = str(self.df.name)
            else:
                df_name = "DataFrame"
            if self.verbose: print(f"[i] Loaded {df_name}{self.df.shape}")
        else:
            print(f"[X] Oops! Missing 'ohlcv' data or index is not datetime ordered.\n")
            return None

        self._validate_mpf_kwargs(**kwargs)
        self._validate_chart_kwargs(**kwargs)
        self._validate_ta_strategy(strategy)

        # Build TA and Plot
        self.df.ta.strategy(self.strategy, verbose=self.verbose)
        self._plot(**kwargs)

    def _validate_ta_strategy(self, strategy):
        if strategy is not None or isinstance(strategy, ta.Strategy):
            self.strategy = strategy
        elif len(self.strategy_ta) > 0:
            print(f"[+] Strategy: {self.strategy_name}")
        else:
            self.strategy = ta.CommonStrategy        

    def _validate_chart_kwargs(self, **kwargs):
        """Chart Settings"""
        self.config = {}
        self.config["last"] = kwargs.pop("last", recent_bars(self.df))
        self.config["rpad"] = kwargs.pop("rpad", 10)
        self.config["title"] = kwargs.pop("title", "Asset")
        self.config["volume"] = kwargs.pop("volume", True)

    def _validate_mpf_kwargs(self, **kwargs):
        # mpf global chart settings
        default_chart = mpf.available_styles()[-1]
        default_mpf_width = {
            'candle_linewidth': 0.6,
            'candle_width': 0.525,
            'volume_width': 0.525
        }
        mpfchart = {}

        mpf_style = kwargs.pop("style", "")
        if mpf_style == "" or mpf_style.lower() == "random":
            mpf_styles = mpf.available_styles()
            mpfchart["style"] = mpf_styles[rnd.randrange(len(mpf_styles))]
        elif mpf_style.lower() in mpf.available_styles():
            mpfchart["style"] = mpf_style

        mpfchart["figsize"] = kwargs.pop("figsize", (12, 10))
        mpfchart["non_trading"] = kwargs.pop("nontrading", False)
        mpfchart["rc"] = kwargs.pop("rc", {'figure.facecolor': '#EDEDED'})
        mpfchart["plot_ratios"] = kwargs.pop("plot_ratios", (12, 1.7))
        mpfchart["scale_padding"] = kwargs.pop("scale_padding", {'left': 1, 'top': 4, 'right': 1, 'bottom': 1})
        mpfchart["tight_layout"] = kwargs.pop("tight_layout", True)
        mpfchart["type"] = kwargs.pop("type", "candle")
        mpfchart["width_config"] = kwargs.pop("width_config", default_mpf_width)
        mpfchart["xrotation"] = kwargs.pop("xrotation", 15)
        
        self.mpfchart = mpfchart

    def _attribution(self):
        print(f"\nPandas v: {pd.__version__} [pip install pandas] https://github.com/pandas-dev/pandas")
        print(f"Data from AlphaVantage v: 1.0.19 [pip install alphaVantage-api] http://www.alphavantage.co https://github.com/twopirllc/AlphaVantageAPI")
        print(f"Technical Analysis with Pandas TA v: {ta.version} [pip install pandas_ta] https://github.com/twopirllc/pandas-ta")
        print(f"Charts by Matplotlib Finance v: {mpf.__version__} [pip install mplfinance] https://github.com/matplotlib/mplfinance\n")

    def _right_pad_df(self, rpad: int, delta_unit: str = "D", range_freq: str = "B"):
        if rpad > 0:
            dfpad = self.df[-rpad:].copy()
            dfpad.iloc[:,:] = np.NaN

            df_frequency = self.df.index.value_counts().mode()[0] # Most common frequency
            freq_delta = pd.Timedelta(df_frequency, unit=delta_unit)
            new_dr = pd.date_range(start=self.df.index[-1] + freq_delta, periods=rpad, freq=range_freq)
            dfpad.index = new_dr # Update the padded index with new dates
#             self.df = self.df.append(dfpad)
            self.df = pd.concat([self.df, dfpad], ignore_index = False)
        
            
    def _plot(self, **kwargs):
        if not isinstance(self.mpfchart["plot_ratios"], tuple):
            print(f"[X] plot_ratios must be a tuple")
            return

        # Override Chart Title Option
        chart_title = self.config["title"]
        if "title" in kwargs and isinstance(kwargs["title"], str):
            chart_title = kwargs.pop("title")

        # Override Right Bar Padding Option
        rpad = self.config["rpad"]
        if "rpad" in kwargs and kwargs["rpad"] > 0:
            rpad = int(kwargs["rpad"])

        def cpanel():
            return len(self.mpfchart["plot_ratios"])

        # Last Second Default TA Indicators
        linreg = kwargs.pop("linreg", False)
        linreg_name = self.df.ta.linreg(append=True).name if linreg else ""

        midpoint = kwargs.pop("midpoint", False)
        midpoint_name = self.df.ta.midpoint(append=True).name if midpoint else ""

        ohlc4 = kwargs.pop("ohlc4", False)
        ohlc4_name = self.df.ta.ohlc4(append=True).name if ohlc4 else ""

        clr = kwargs.pop("clr", False)
        clr_name = self.df.ta.log_return(cumulative=True, append=True).name if clr else ""

        rsi = kwargs.pop("rsi", False)
        rsi_length = kwargs.pop("rsi_length", None)
        if isinstance(rsi_length, int) and rsi_length > 1:
            rsi_name = self.df.ta.rsi(length=rsi_length, append=True).name
        elif rsi:
            rsi_name = self.df.ta.rsi(append=True).name
        else: rsi_name = ""
            
        zscore = kwargs.pop("zscore", False)
        zscore_length = kwargs.pop("zscore_length", None)
        if isinstance(zscore_length, int) and zscore_length > 1:
            zs_name = self.df.ta.zscore(length=zscore_length, append=True).name
        elif zscore:
            zs_name = self.df.ta.zscore(append=True).name
        else: zs_name = ""

        macd = kwargs.pop("macd", False)
        macd_name = ""
        if macd:
            macds = self.df.ta.macd(append=True)
            macd_name = macds.name

        squeeze = kwargs.pop("squeeze", False)
        lazybear = kwargs.pop("lazybear", False)
        squeeze_name = ""
        if squeeze:
            squeezes = self.df.ta.squeeze(lazybear=lazybear, detailed=True, append=True)
            squeeze_name = squeezes.name

        ama = kwargs.pop("archermas", False)
        ama_name = ""
        if ama:
            amas = self.df.ta.amat(append=True)
            ama_name = amas.name

        aobv = kwargs.pop("archerobv", False)
        aobv_name = ""
        if aobv:
            aobvs = self.df.ta.aobv(append=True)
            aobv_name = aobvs.name

        # Pad and trim Chart
        self._right_pad_df(rpad)
        mpfdf = self.df.tail(self.config["last"]).copy()
        mpfdf_columns = list(self.df.columns)
        
        tsig = kwargs.pop("tsignals", False)
        if tsig:
            # Long Trend requires Series Comparison (<=. <, = >, >=)
            # or Trade Logic that yields trends in binary.
            default_long = mpfdf["SMA_10"] > mpfdf["SMA_20"]
            long_trend = kwargs.pop("long_trend", default_long)
            if not isinstance(long_trend, pd.Series):
                raise(f"[X] Must be a Series that has boolean values or values of 0s and 1s")
            mpfdf.ta.percent_return(append=True)
            mpfdf.ta.tsignals(long_trend, append=True)
            buys = np.where(mpfdf.TS_Entries > 0, 1, np.nan)
            sells = np.where(mpfdf.TS_Exits > 0, 1, np.nan)
            mpfdf["ACTRET_1"] = mpfdf.TS_Trends * mpfdf.PCTRET_1

        # BEGIN: Custom TA Plots and Panels
        # Modify the area below 
        taplots = [] # Holds all the additional plots

        # Panel 0: Price Overlay
        if linreg_name in mpfdf_columns:
            taplots += [mpf.make_addplot(mpfdf[linreg_name], type=kwargs.pop("linreg_type", "line"), color=kwargs.pop("linreg_color", "black"), linestyle="-.", width=1.2, panel=0)]

        if midpoint_name in mpfdf_columns:
            taplots += [mpf.make_addplot(mpfdf[midpoint_name], type=kwargs.pop("midpoint_type", "scatter"), color=kwargs.pop("midpoint_color", "fuchsia"), width=0.4, panel=0)]

        if ohlc4_name in mpfdf_columns:
            taplots += [mpf.make_addplot(mpfdf[ohlc4_name], ylabel=ohlc4_name, type=kwargs.pop("ohlc4_type", "scatter"), color=kwargs.pop("ohlc4_color", "blue"), alpha=0.85, width=0.4, panel=0)]
    
        if self.strategy.name == ta.CommonStrategy.name:
            total_sma = 0 # Check if all the overlap indicators exists before adding plots
            for c in ["SMA_10", "SMA_20", "SMA_50", "SMA_200"]:
                if c in mpfdf_columns: total_sma += 1
                else: print(f"[X] Indicator: {c} missing!")
            if total_sma == 4:
                ta_smas = [
                    mpf.make_addplot(mpfdf["SMA_10"], color="green", width=1.5, panel=0),
                    mpf.make_addplot(mpfdf["SMA_20"], color="orange", width=2, panel=0),
                    mpf.make_addplot(mpfdf["SMA_50"], color="red", width=2, panel=0),
                    mpf.make_addplot(mpfdf["SMA_200"], color="maroon", width=3, panel=0),
                ]
                taplots += ta_smas

        if tsig:
            taplots += [
                mpf.make_addplot(0.985 * mpfdf.close * buys, type="scatter", marker="^", markersize=26, color="blue", panel=0),
                mpf.make_addplot(1.015 * mpfdf.close * sells, type="scatter", marker="v", markersize=26, color="fuchsia", panel=0),
            ]
                
        if len(ama_name):
            amat_sr_ = mpfdf[amas.columns[-1]][mpfdf[amas.columns[-1]] > 0]
            amat_sr = amat_sr_.index.to_list()
        else:
            amat_sr = None

        # Panel 1: If volume=True, the add the VOL MA. Since we know there is only one, we immediately pop it.
        if self.config["volume"]:
            volma = [x for x in list(self.df.columns) if x.startswith("VOL_")].pop()
            max_vol = mpfdf["volume"].max()
             # Volume axis
            ta_volume = [mpf.make_addplot(mpfdf[volma], color="black", width=1.2, panel=1, ylim=(-.2 * max_vol, 1.5 * max_vol))]
            taplots += ta_volume

        # Panels 2 - 9
        common_plot_ratio = (3,)

        if len(aobv_name):
            _p = kwargs.pop("aobv_percenty", 0.2)
            aobv_ylim = ta_ylim(mpfdf[aobvs.columns[0]], _p)
            taplots += [
                mpf.make_addplot(mpfdf[aobvs.columns[0]], ylabel=aobv_name, color="black", width=1.5, panel=cpanel(), ylim=aobv_ylim),
                mpf.make_addplot(mpfdf[aobvs.columns[2]], color="silver", width=1, panel=cpanel(), ylim=aobv_ylim),
                mpf.make_addplot(mpfdf[aobvs.columns[3]], color="green", width=1, panel=cpanel(), ylim=aobv_ylim),
                mpf.make_addplot(mpfdf[aobvs.columns[4]], color="red", width=1.2, panel=cpanel(), ylim=aobv_ylim),
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

        if clr_name in mpfdf_columns:
            _p = kwargs.pop("clr_percenty", 0.1)
            clr_ylim = ta_ylim(mpfdf[clr_name], _p)

            taplots += [mpf.make_addplot(mpfdf[clr_name], ylabel=clr_name, color="black", width=1.5, panel=cpanel(), ylim=clr_ylim)]
            if (1 - _p) * mpfdf[clr_name].min() < 0 and (1 + _p) * mpfdf[clr_name].max() > 0:
                taplots += [mpf.make_addplot(mpfdf["0"], color="gray", width=1.2, panel=cpanel(), ylim=clr_ylim)]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

        if rsi_name in mpfdf_columns:
            rsi_ylim = (0, 100)
            taplots += [
                mpf.make_addplot(mpfdf[rsi_name], ylabel=rsi_name, color=kwargs.pop("rsi_color", "black"), width=1.5, panel=cpanel(), ylim=rsi_ylim),
                mpf.make_addplot(mpfdf["20"], color="green", width=1, panel=cpanel(), ylim=rsi_ylim),
                mpf.make_addplot(mpfdf["50"], color="gray", width=0.8, panel=cpanel(), ylim=rsi_ylim),
                mpf.make_addplot(mpfdf["80"], color="red", width=1, panel=cpanel(), ylim=rsi_ylim),
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel
        
        if macd_name in mpfdf_columns:
            _p = kwargs.pop("macd_percenty", 0.15)
            macd_ylim = ta_ylim(mpfdf[macd_name], _p)
            taplots += [
                mpf.make_addplot(mpfdf[macd_name], ylabel=macd_name, color="black", width=1.5, panel=cpanel()),#, ylim=macd_ylim),
                mpf.make_addplot(mpfdf[macds.columns[-1]], color="blue", width=1.1, panel=cpanel()),#, ylim=macd_ylim),
                mpf.make_addplot(mpfdf[macds.columns[1]], type="bar", alpha=0.8, color="dimgray", width=0.8, panel=cpanel()),#, ylim=macd_ylim),
                mpf.make_addplot(mpfdf["0"], color="black", width=1.2, panel=cpanel()),#, ylim=macd_ylim),
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel            

        if zs_name in mpfdf_columns:
            _p = kwargs.pop("zascore_percenty", 0.2)
            zs_ylim = ta_ylim(mpfdf[zs_name], _p)
            taplots += [
                mpf.make_addplot(mpfdf[zs_name], ylabel=zs_name, color="black", width=1.5, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["-3"], color="red", width=1.2, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["-2"], color="orange", width=1, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["-1"], color="silver", width=1, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["0"], color="black", width=1.2, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["1"], color="silver", width=1, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["2"], color="orange", width=1, panel=cpanel(), ylim=zs_ylim),
                mpf.make_addplot(mpfdf["3"], color="red", width=1.2, panel=cpanel(), ylim=zs_ylim)
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

        if squeeze_name in mpfdf_columns:
            _p = kwargs.pop("squeeze_percenty", 0.6)
            sqz_ylim = ta_ylim(mpfdf[squeeze_name], _p)
            taplots += [
                mpf.make_addplot(mpfdf[squeezes.columns[-4]], type="bar", color="lime", alpha=0.65, width=0.8, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf[squeezes.columns[-3]], type="bar", color="green", alpha=0.65, width=0.8, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf[squeezes.columns[-2]], type="bar", color="maroon", alpha=0.65, width=0.8, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf[squeezes.columns[-1]], type="bar", color="red", alpha=0.65, width=0.8, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf["0"], color="black", width=1.2, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf[squeezes.columns[4]], ylabel=squeeze_name, color="green", width=2, panel=cpanel(), ylim=sqz_ylim),
                mpf.make_addplot(mpfdf[squeezes.columns[5]], color="red", width=1.8, panel=cpanel(), ylim=sqz_ylim),
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

        if tsig:
            _p = kwargs.pop("tsig_percenty", 0.23)
            treturn_ylim = ta_ylim(mpfdf["ACTRET_1"], _p)
            taplots += [
                mpf.make_addplot(mpfdf["ACTRET_1"], ylabel="Active % Return", type="bar", color="green", alpha=0.45, width=0.8, panel=cpanel(), ylim=treturn_ylim),
                mpf.make_addplot(pd.Series(mpfdf["ACTRET_1"].mean(), index=mpfdf["ACTRET_1"].index), color="blue", width=1, panel=cpanel(), ylim=treturn_ylim),
                mpf.make_addplot(mpfdf["0"], color="black", width=1, panel=cpanel(), ylim=treturn_ylim),
            ]
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

            _p = kwargs.pop("cstreturn_percenty", 0.58)
            mpfdf["CUMACTRET_1"] = mpfdf["ACTRET_1"].cumsum()
            cumactret_ylim = ta_ylim(mpfdf["CUMACTRET_1"], _p)
            taplots += [
                mpf.make_addplot(mpfdf["CUMACTRET_1"], ylabel="Cum Trend Return", type="bar", color="silver", alpha=0.45, width=1, panel=cpanel(), ylim=cumactret_ylim),
                mpf.make_addplot(0.9 * buys * mpfdf["CUMACTRET_1"], type="scatter", marker="^", markersize=14, color="green", panel=cpanel(), ylim=cumactret_ylim),
                mpf.make_addplot(1.1 * sells * mpfdf["CUMACTRET_1"], type="scatter", marker="v", markersize=14, color="red", panel=cpanel(), ylim=cumactret_ylim),
                mpf.make_addplot(mpfdf["0"], color="black", width=1, panel=cpanel(), ylim=cumactret_ylim),
            ]            
            self.mpfchart["plot_ratios"] += common_plot_ratio # Required to add a new Panel

        # END: Custom TA Plots and Panels
        
        if self.verbose:
            additional_ta = []
            chart_title  = f"{chart_title} [{self.strategy.name}] (last {self.config['last']} bars)"
            chart_title += f"\nSince {mpfdf.index[0]} till {mpfdf.index[-1]}"
            if len(linreg_name) > 0: additional_ta.append(linreg_name)
            if len(midpoint_name) > 0: additional_ta.append(midpoint_name)
            if len(ohlc4_name) > 0: additional_ta.append(ohlc4_name)
            if len(additional_ta) > 0:
                chart_title += f"\nIncluding: {', '.join(additional_ta)}"

        if amat_sr:
            vlines_ = dict(vlines=amat_sr, alpha=0.1, colors="red")
        else:
            # Hidden because vlines needs valid arguments even if None 
            vlines_ = dict(vlines=mpfdf.index[0], alpha=0, colors="white")

        # Create Final Plot
        mpf.plot(mpfdf,
            title=chart_title,
            type=self.mpfchart["type"],
            style=self.mpfchart["style"],
            datetime_format="%-m/%-d/%Y %m-%d %H",
            volume=self.config["volume"],
            figsize=self.mpfchart["figsize"],
            tight_layout=self.mpfchart["tight_layout"],
            scale_padding=self.mpfchart["scale_padding"],
            panel_ratios=self.mpfchart["plot_ratios"], # This key needs to be update above if adding more panels
            xrotation=self.mpfchart["xrotation"],
            update_width_config=self.mpfchart["width_config"],
            show_nontrading=self.mpfchart["non_trading"],
            vlines=vlines_,
            addplot=taplots
        )

        
        self._attribution()
        
def recent_bars(df, tf: str = "1y"):
    # All Data: 0, Last Four Years: 0.25, Last Two Years: 0.5, This Year: 1, Last Half Year: 2, Last Quarter: 4
    yearly_divisor = {"all": 0, "10y": 0.1, "5y": 0.2, "4y": 0.25, "3y": 1./3, "2y": 0.5, "1y": 1, "6mo": 2, "3mo": 4, "10D": 1}
    if (tf == '10D'):
        return 10*26
    yd = yearly_divisor[tf] if tf in yearly_divisor.keys() else 0
    return int(ta.RATE["TRADING_DAYS_PER_YEAR"] / yd) if yd > 0 else df.shape[0]

DEBUG=0

    
def trade_asset(asset, trendy):
    entries = trendy.TS_Entries * asset.close
    entries = entries[~np.isclose(entries, 0)]
    entries.dropna(inplace=True)
    entries.name = "Entry"

    exits = trendy.TS_Exits * asset.close
    exits = exits[~np.isclose(exits, 0)]
    exits.dropna(inplace=True)
    exits.name = "Exit"

    total_trades = trendy.TS_Trades.abs().sum()
    rt_trades = int(trendy.TS_Trades.abs().sum() // 2)

    all_trades = trendy.TS_Trades.copy().fillna(0)
    all_trades = all_trades[all_trades != 0]

    trades = pd.DataFrame({
        "Signal": all_trades,
        entries.name: entries,
        exits.name: exits
    })

    # Show some stats if there is an active trade (when there is an odd number of round trip trades)
    if total_trades % 2 != 0:
        unrealized_pnl = asset.close.iloc[-1] - entries.iloc[-1]
        unrealized_pnl_pct_change = 100 * ((asset.close.iloc[-1] / entries.iloc[-1]) - 1)
        print("Current Trade:")
        print(f"Price Entry | Last:\t{entries.iloc[-1]:.4f} | {asset.close.iloc[-1]:.4f}")
        print(f"Unrealized PnL | %:\t{unrealized_pnl:.4f} | {unrealized_pnl_pct_change:.4f}%")
    print(f"\nTrades Total | Round Trip:\t{total_trades} | {rt_trades}")
    print(f"Trade Coverage: {100 * asset.TS_Trends.sum() / asset.shape[0]:.2f}%")

    if DEBUG==1:
        print(trades)
    return asset, trades.tail(1)

def retrieve_data(tickers, tf='D', stratName = None):
    watch = Watchlist(tickers, tf=tf, ds_name="yahoo", timed=True)
    # watch.strategy = ta.CommonStrategy # If you have a Custom Strategy, you can use it here.
    momo_bands_sma_ta = [
        {"kind":"sma", "length": 50},
        {"kind":"sma", "length": 100},
        {"kind":"sma", "length": 200},
        {"kind":"sma", "length": 250},
        {"kind":"bbands", "length": 20, "ddof": 0},
        {"kind":"macd"},
        {"kind":"rsi"},
        {"kind":"ppo"},
        {"kind":"log_return", "cumulative": True},
        {"kind":"sma", "close": "CUMLOGRET_1", "length": 5, "suffix": "CUMLOGRET"},
    ]
    momo_bands_sma_strategy = ta.Strategy(
        "Momo, Bands and SMAs and Cumulative Log Returns", # name
        momo_bands_sma_ta, # ta
        "MACD and RSI Momo with BBANDS and SMAs 50 & 200 and Cumulative Log Returns" # description
    )
    if stratName is None:
        watch.strategy = momo_bands_sma_ta
    else:
        watch.strategy = stratName
    watch.load(tickers, analyze=True, verbose=False)
    
    return watch

# def read_assets(tickers, start_data, end_date):
#     assets = vbt.YFData.download(tickers, 
#                       start = start_date, 
#                       end = end_date, 
#                       interval="D",                                    
#                       missing_index="drop")
#     return assets

def add_jc_data(asset):
#     asset = watch.data[ticker]
    sma=50
    rolling_window = (asset['Close']-asset['SMA_50']).rolling(window=sma)
    rolling_rank = rolling_window.apply(lambda x: pd.Series(x).rank().iloc[-1])
    asset['rk50'] = rolling_rank
    sma=100
    rolling_window = (asset['Close']-asset['SMA_100']).rolling(window=sma)
    rolling_rank = rolling_window.apply(lambda x: pd.Series(x).rank().iloc[-1])
    
    asset['rk100'] = rolling_rank
    asset['cmb_rk'] = (asset['rk50'] + asset['rk100'])*100/150
    
    sma=250
    rolling_window = asset['Close'].rolling(window=sma)
    rolling_rank = rolling_window.apply(lambda x: pd.Series(x).rank().iloc[-1])
    asset['rk250'] = rolling_rank
    sma=100
    rolling_window = asset['Close'].rolling(window=sma)
    rolling_rank = rolling_window.apply(lambda x: pd.Series(x).rank().iloc[-1])
    asset['rkprice'] = rolling_rank
    asset['deltaclose'] = asset['Close'].diff()    
    return asset

def dtmask(df: pd.DataFrame, start: datetime, end: datetime):
    df['Datetime'] = pd.to_datetime(df.index, utc=True)
#     print(df['Datetime'])
    if not df.ta.datetime_ordered:
        df = df.set_index(pd.DatetimeIndex(df['Datetime']))    
    return df.loc[(df.index >= pd.to_datetime(start, utc=True)) & (df.index <= pd.to_datetime(end, utc=True)), :].copy()

def dtmonth(df: pd.DataFrame, start , end):
    df['Day'] = df.index.strftime('%d').astype(int)
    df['Month'] = df.index.strftime('%m').astype(int)
    return df.loc[df.index.strftime('%Y-%m-%d').isin(start) | df.index.strftime('%Y-%m-%d').isin(end) , :].copy()


def check_cross(df, sma = 200):
    if sma==50:
        sma_value = df['SMA_50']
    elif sma==100:
        sma_value = df['SMA_100']
    elif sma==250:
        sma_value = df['SMA_250']
    else:
        sma_value = df['SMA_200']
        
    df['Day-1'] = df['Day'].shift(1)
    df['close-1'] = df['Close'].shift(1)

    df['cross'] =  np.where((df['Close'] > sma_value) & (df['close-1'] < sma_value) & (df['Day'] > 15) & (df['Day-1'] < 15), 1, 0)
    df['cross'] =  np.where((df['Close'] < sma_value) & (df['close-1'] > sma_value) & (df['Day'] > 15) & (df['Day-1']< 15), -1, df['cross'])
    

def process_data(watch, ticker, tf, duration="10y"):    
    if DEBUG==1:
        print(f"{ticker} {watch.data[ticker].shape}\nColumns: {', '.join(list(watch.data[ticker].columns))}")

    #
    # trim data, store to 'asset'
    #
    asset = watch.data[ticker]
    recent = recent_bars(asset, duration)
    asset.columns = asset.columns.str.lower()
    asset.drop(columns=["dividends", "split"], errors="ignore", inplace=True)
    asset = asset.copy().tail(recent)
    asset.name = ticker
#     if DEBUG==1:
#         print(asset)
    
    # Example Long Trends
    # long = ta.sma(asset.close, 50) > ta.sma(asset.close, 200) # SMA(50) > SMA(200) "Golden/Death Cross"
    # long = ta.sma(asset.close, 10) > ta.sma(asset.close, 20) # SMA(10) > SMA(20)
    long = ta.ema(asset.close, 5) > ta.ema(asset.close, 34) # EMA(8) > EMA(21)
#     long = ta.increasing(ta.ema(asset.close, 50))
#     long = ta.macd(asset.close).iloc[:,1] > 0 # MACD Histogram is positive
    # long = ta.amat(asset.close, 50, 200).AMATe_LR_2  # Long Run of AMAT(50, 200) with lookback of 2 bars

    # long &= ta.increasing(ta.ema(asset.close, 50), 2) # Uncomment for further long restrictions, in this case when EMA(50) is increasing/sloping upwards
    # long = 1 - long # uncomment to create a short signal of the trend

    ###############################################################################################
    # PPO long
    # AND [Daily SMA(20,Daily Volume) > 40000] 
    # AND [Daily SMA(60,Daily Close) > 20] 

    # AND [Daily Close > Daily SMA(200,Daily Close)] 
    # AND [Yesterday's Daily PPO Line(12,26,9,Daily Close) < Daily PPO Signal(12,26,9,Daily Close)] 
    # AND [Daily PPO Line(12,26,9,Daily Close) > Daily PPO Signal(12,26,9,Daily Close)] 
    # AND [Daily PPO Line(12,26,9,Daily Close) < 0]
    #
    ###############################################################################################
    asset.ta.ema(length=5, sma=False, append=True)
    asset.ta.ema(length=34, sma=False, append=True)
    asset.ta.ema(length=50, sma=False, append=True)
    asset.ta.percent_return(append=True)
    print("TA Columns Added:")
    if DEBUG==1:
        print(asset[asset.columns[5:]].tail())
        
    trendy = asset.ta.tsignals(long, asbool=False, append=True)
    if DEBUG==1:
        print(trendy.tail())
    
    return asset, trendy

def display_asset(asset, tf='D', duration="1y"):
    extime = ta.get_time(to_string=True)
    recent = recent_bars(asset, duration)
    
    first_date, last_date = asset.index[0], asset.index[-1]
    f_date = f"{first_date.day_name()} {first_date.month}-{first_date.day}-{first_date.year}"
    l_date = f"{last_date.day_name()} {last_date.month}-{last_date.day}-{last_date.year}"
    last_ohlcv = f"Last OHLCV: ({asset.iloc[-1].open:.4f}, {asset.iloc[-1].high:.4f}, {asset.iloc[-1].low:.4f}, {asset.iloc[-1].close:.4f}, {int(asset.iloc[-1].volume)})"
    ptitle = f"\n{asset.name} [{tf} for {duration}({recent} bars)] from {f_date} to {l_date}\n{last_ohlcv}\n{extime}"

    # chart = asset["close"] #asset[["close", "SMA_10", "SMA_20", "SMA_50", "SMA_200"]]
    # chart = asset[["close", "SMA_10", "SMA_20"]]
    chart = asset[["close", "EMA_5", "EMA_34", "EMA_50"]]
#     obv_addplot = mpf.make_addplot(asset["PPO_12_26_9"], width=2, panel=2, ylabel="OBV")
    chart.index = chart.index.astype(str)        

    chart.plot(figsize=(16, 10), color=colors("BkGrOrRd"), title=ptitle, grid=True)

def display_trend_1(trendy):
    long_trend = trendy.TS_Trends
    short_trend = 1 - long_trend

    long_trend.index = long_trend.index.astype(str)
    short_trend.index = short_trend.index.astype(str)
    
    long_trend.plot(figsize=(16, 0.85), kind="area", stacked=True, color=colors()[0], alpha=0.25) # Green Area
    short_trend.plot(figsize=(16, 0.85), kind="area", stacked=True, color=colors()[1], alpha=0.25) # Red Area
    
def display_trend_2(trendy):
    trendy.TS_Trades.index = trendy.TS_Trades.index.astype(str)
    trendy.TS_Trades.plot(figsize=(16, 1.5), color=colors("BkBl")[0], grid=True)

def display_return_1(asset, trendy):
    asset["ACTRET_1"] = trendy.TS_Trends.shift(1) * asset.PCTRET_1
#     asset[["PCTRET_1", "ACTRET_1"]].index = asset[["PCTRET_1", "ACTRET_1"]].index.astype(str)
    
    asset[["PCTRET_1", "ACTRET_1"]].plot(figsize=(16, 3), color=colors("GyOr"), alpha=1, grid=True).axhline(0, color="black")    

def display_return_2(asset, trendy):
    asset["ACTRET_1"] = trendy.TS_Trends.shift(1) * asset.PCTRET_1
#     asset["ACTRET_1"].index = asset["ACTRET_1"].index.astype(str)
#     asset["PCTRET_1"].index = asset["PCTRET_1"].index.astype(str)    
    ((asset[["PCTRET_1", "ACTRET_1"]] + 1).cumprod() - 1).plot(figsize=(16, 3), kind="area", stacked=False, color=colors("GyOr"), title="B&H vs. Cum. Active Returns", alpha=.4, grid=True).axhline(0, color="black")
    
def combine_stats(pf: vbt.portfolio.base.Portfolio, ticker: str, strategy: str, mode: int = 0):
    header = pd.Series({
        "Run Time": ta.get_time(full=False, to_string=True),
        "Mode": "LIVE" if mode else "TEST",
        "Strategy": strategy,
        "Direction": vbt.settings.portfolio["signal_direction"],
        "Symbol": ticker.upper(),
        "Fees [%]": 100 * vbt.settings.portfolio["fees"],
        "Slippage [%]": 100 * vbt.settings.portfolio["slippage"],
        "Accumulate": vbt.settings.portfolio["accumulate"],
    })
    rstats = pf.returns_stats().dropna(axis=0).T
    stats = pf.stats().dropna(axis=0).T
    joint = pd.concat([header, stats, rstats])
    return joint[~joint.index.duplicated(keep="first")]

def earliest_common_index(d: dict):
    """Returns index of the earliest common index of all DataFrames in the dict"""
    min_date = None
    for df in d.values():
        if min_date is None:
            min_date = df.index[0]
        elif min_date < df.index[0]:
            min_date = df.index[0]
    return min_date

def dl(tickers: list, same_start: bool = False, **kwargs):
    if isinstance(tickers, str):
        tickers = [tickers]
    
    if not isinstance(tickers, list) or len(tickers) == 0:
        print("Must be a non-empty list of tickers or symbols")
        return

    if "limit" in kwargs and kwargs["limit"] and len(tickers) > kwargs["limit"]:
        from itertools import islice            
        tickers = list(islice(tickers, kwargs["limit"]))
        print(f"[!] Too many assets to compare. Using the first {kwargs['limit']}: {', '.join(tickers)}")

    print(f"[i] Downloading: {', '.join(tickers)}")

    received = {}
    if len(tickers):
        _df = pd.DataFrame()
        for ticker in tickers:
            received[ticker] = _df.ta.ticker(ticker, **kwargs)
            print(f"[+] {ticker}{received[ticker].shape} {ta.get_time(full=False, to_string=True)}")
    
    if same_start and len(tickers) > 1:
        earliestci = earliest_common_index(received)
        print(f"[i] Earliest Common Date: {earliestci}")
        result = {ticker:df[df.index > earliestci].copy() for ticker,df in received.items()}
    else:
        result = received
    print(f"[*] Download Complete\n")
    return result

def dtmask(df: pd.DataFrame, start: datetime, end: datetime):
    df['Datetime'] = pd.to_datetime(df.index, utc=True)
#     print(df['Datetime'])
    if not df.ta.datetime_ordered:
        df = df.set_index(pd.DatetimeIndex(df['Datetime']))    
    return df.loc[(df.index >= pd.to_datetime(start, utc=True)) & (df.index <= pd.to_datetime(end, utc=True)), :].copy()

def dtmonth(df: pd.DataFrame, start , end):
    df['Day'] = df.index.strftime('%d').astype(int)
    df['Month'] = df.index.strftime('%m').astype(int)
    return df.loc[df.index.strftime('%Y-%m-%d').isin(start) | df.index.strftime('%Y-%m-%d').isin(end) , :].copy()

def show_data(d: dict):
    [print(f"{t}[{df.index[0]} - {df.index[-1]}]: {df.shape} {df.ta.time_range:.2f} years") for t,df in d.items()]
    
def trade_table(pf: vbt.portfolio.base.Portfolio, k: int = 1, total_fees: bool = False):
    if not isinstance(pf, vbt.portfolio.base.Portfolio): return
    k = int(k) if isinstance(k, int) and k > 0 else 1

    df = pf.trades.records[["status", "direction", "size", "entry_price", "exit_price", "return", "pnl"]]
    if total_fees:
        df["total_fees"] = df["entry_fees"] + df["exit_fees"]
#     df.to_excel("trade_udow_2016to2023_exits.xlsx")
    print(f"\nLast {k} of {df.shape[0]} Trades\n{df.tail(k)}\n")    
    