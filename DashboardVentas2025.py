import pandas as pd
import plotly.express as px
import streamlit as st

# Function to load data
@st.cache_data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        st.error(f"Error: El archivo '{file_path}' no se encontró. Asegúrate de que está en la misma carpeta que la aplicación Streamlit.")
        st.stop() # Detiene la ejecución si el archivo no se encuentra
    except Exception as e:
        st.error(f"Error al cargar los datos: {e}")
        st.stop()

# Function to get top products by sales
def get_top_products_by_sales(df, n=5):
    # Ensure 'Sales' column exists and is numeric
    if 'Sales' in df.columns:
        df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce').fillna(0) # Convert to numeric, handle errors
    else:
        st.warning("La columna 'Sales' no se encuentra en el archivo de datos.")
        return pd.DataFrame()

    product_sales = df.groupby('Product Name')[['Sales', 'Quantity']].sum().reset_index()
    top_products = product_sales.sort_values(by='Sales', ascending=False).head(n)
    return top_products

# Function to get top products by profit
def get_top_products_by_profit(df, n=5):
    # Ensure 'Profit' column exists and is numeric
    if 'Profit' in df.columns:
        df['Profit'] = pd.to_numeric(df['Profit'], errors='coerce').fillna(0) # Convert to numeric, handle errors
    else:
        st.warning("La columna 'Profit' no se encuentra en el archivo de datos.")
        return pd.DataFrame()

    product_sales_profit = df.groupby('Product Name')[['Sales', 'Profit']].sum().reset_index()
    top_profit_products = product_sales_profit.sort_values(by='Profit', ascending=False).head(n)
    return top_profit_products

# Function to create bar chart for sales
def create_sales_bar_chart(df_top_products, title):
    if not df_top_products.empty:
        fig = px.bar(df_top_products, x='Product Name', y='Sales', title=title, text_auto=True)
        fig.update_layout(xaxis=dict(tickangle=-45, automargin=True, tickfont=dict(size=10)), yaxis=dict(title='Sales'), xaxis_title='Product Name')
        return fig
    else:
        return None

# Function to create bar chart for profit
def create_profit_bar_chart(df_top_profit_products, title):
    if not df_top_profit_products.empty:
        fig = px.bar(df_top_profit_products, x='Product Name', y='Profit', title=title)
        fig.update_layout(xaxis=dict(tickangle=-45, automargin=True, tickfont=dict(size=10)), yaxis=dict(title='Profit'), xaxis_title='Product Name')
        return fig
    else:
        return None

# Streamlit App
def main():
    st.title("Análisis de Ventas y Ganancias de Productos")

    file_path = "SalidaVentas.xlsx"
    df = load_data(file_path)

    if df is None: # If load_data failed and stopped, this will be skipped
        return

    st.sidebar.header("Filtros")

    region_options = ['Todas'] + list(df['Region'].unique())
    selected_region = st.sidebar.selectbox("Selecciona una Región", region_options)

    filtered_df = df.copy() # Use .copy() to avoid SettingWithCopyWarning

    if selected_region != 'Todas':
        filtered_df = filtered_df[filtered_df['Region'] == selected_region]

    # Check if 'State' column exists before trying to filter by it
    if 'State' in filtered_df.columns and not filtered_df['State'].empty:
        state_options = ['Todos'] + list(filtered_df['State'].unique())
        selected_state = st.sidebar.selectbox("Selecciona un Estado", state_options)

        if selected_state != 'Todos':
            filtered_df = filtered_df[filtered_df['State'] == selected_state]
    else:
        selected_state = 'No Disponible' # Default value if 'State' column is missing or empty

    st.subheader(f"Análisis de Ventas y Ganancias para: {selected_region} - {selected_state}")

    # Display Top Products by Sales
    st.subheader("Top 5 Productos por Ventas")
    top_products_sales = get_top_products_by_sales(filtered_df)
    if not top_products_sales.empty:
        fig_sales = create_sales_bar_chart(top_products_sales, "Top 5 Productos por Ventas")
        if fig_sales:
            st.plotly_chart(fig_sales)
        else:
            st.info("No hay datos de ventas para mostrar en los productos principales.")
    else:
        st.info("No hay datos suficientes para mostrar los productos principales por ventas.")

    # Display Top Products by Profit
    st.subheader("Top 5 Productos por Ganancias")
    top_products_profit = get_top_products_by_profit(filtered_df)
    if not top_products_profit.empty:
        fig_profit = create_profit_bar_chart(top_products_profit, "Top 5 Productos por Ganancias")
        if fig_profit:
            st.plotly_chart(fig_profit)
        else:
            st.info("No hay datos de ganancias para mostrar en los productos principales.")
    else:
        st.info("No hay datos suficientes para mostrar los productos principales por ganancias.")

# Run the Streamlit app
if __name__ == "__main__":
    main()
