# Generated by Django 4.2.7 on 2024-10-23 16:46

from django.db import migrations


class Migration(migrations.Migration):

    dependencies = [
        ('patrimoine', '0002_rename_id_region_region_id_and_more'),
    ]

    operations = [
        migrations.RenameField(
            model_name='region',
            old_name='region_id',
            new_name='id',
        ),
        migrations.RenameField(
            model_name='service',
            old_name='service_id',
            new_name='id',
        ),
    ]
