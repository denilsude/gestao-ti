# Usa uma imagem oficial do Python como base
FROM python:3.11-slim

# Define variáveis de ambiente
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Copia e instala as dependências (primeiro para otimizar o cache)
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copia o restante do código da aplicação
COPY . /app/

# Expõe a porta
EXPOSE 5000

# Comando para iniciar a aplicação Flask
CMD ["python", "app.py"]