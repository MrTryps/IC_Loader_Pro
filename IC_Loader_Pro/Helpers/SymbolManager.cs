using ArcGIS.Core.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading.Tasks;

namespace IC_Loader_Pro.Helpers // Or another appropriate namespace
{
    /// <summary>
    /// A static helper class to manage and cache symbols loaded from the project's style file.
    /// </summary>
    public static class SymbolManager
    {
        // A thread-safe dictionary to cache symbols after they are loaded once.
        private static readonly ConcurrentDictionary<string, CIMSymbol> _symbolCache = new ConcurrentDictionary<string, CIMSymbol>();
        private const string StyleFileName = "ICLoader_Symbols";

        /// <summary>
        /// A generic method to retrieve a symbol from the cache or load it from the .stylx file.
        /// </summary>
        /// <typeparam name="T">The type of CIMSymbol to retrieve (e.g., CIMPolygonSymbol, CIMPointSymbol).</typeparam>
        /// <param name="symbolName">The name (key) of the symbol to find in the style file.</param>
        /// <returns>The requested symbol, or null if not found.</returns>
        public static async Task<T> GetSymbolAsync<T>(string symbolName) where T : CIMSymbol
        {
            // 1. Check the cache first.
            if (_symbolCache.TryGetValue(symbolName, out CIMSymbol symbol))
            {
                return symbol as T;
            }

            // 2. If not in the cache, load it from the style file.
            SymbolStyleItem styleItem = null;
            await QueuedTask.Run(() =>
            {
                var styleProjectItem = Project.Current.GetItems<StyleProjectItem>()
                                              .FirstOrDefault(s => s.Name == StyleFileName);
                if (styleProjectItem == null)
                {
                    // Log error if the .stylx file isn't found
                    System.Diagnostics.Debug.WriteLine($"Style file '{StyleFileName}.stylx' not found.");
                    return;
                }

                try
                {
                    // It uses the synchronous SearchSymbols method and gets the first result.
                    styleItem = styleProjectItem.SearchSymbols(GetStyleItemType<T>(), symbolName)[0];
                    // --------------------------------------------------
                }
                catch (Exception ex)
                {
                    // This will catch the error if the symbol is not found (IndexOutOfRangeException)
                    System.Diagnostics.Debug.WriteLine($"Error finding symbol '{symbolName}' in '{StyleFileName}.stylx'. Exception: {ex.Message}");
                    styleItem = null;
                }
            });

            if (styleItem != null)
            {
                var newSymbol = styleItem.Symbol as T;
                if (newSymbol != null)
                {
                    _symbolCache.TryAdd(symbolName, newSymbol);
                    return newSymbol;
                }
            }

            // If we reach here, the symbol was not found or failed to load.
            return null;
        }

        /// <summary>
        /// Helper to determine the StyleItemType from the generic type parameter.
        /// </summary>
        private static StyleItemType GetStyleItemType<T>() where T : CIMSymbol
        {
            if (typeof(T) == typeof(CIMPointSymbol)) return StyleItemType.PointSymbol;
            if (typeof(T) == typeof(CIMPolygonSymbol)) return StyleItemType.PolygonSymbol;
            if (typeof(T) == typeof(CIMLineSymbol)) return StyleItemType.LineSymbol;
            // Add other types as needed
            return StyleItemType.Unknown;
        }
    }
}