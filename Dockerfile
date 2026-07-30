# Run the dashboard/summary generators without installing Python:
#
#   docker run --rm -v "$PWD/records:/data" ghcr.io/lawrenceleejr/researchgroupinfo \
#       /data -o /data/dashboard.html --title "My Group"
#
#   docker run --rm -v "$PWD/records:/data" --entrypoint python \
#       ghcr.io/lawrenceleejr/researchgroupinfo /app/generate_summaries.py /data -o /data
FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY generate_dashboard.py generate_summaries.py build_template.py make_sample_data.py ./
COPY template/ template/

WORKDIR /data
ENTRYPOINT ["python", "/app/generate_dashboard.py"]
CMD ["--help"]
