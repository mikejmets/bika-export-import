import argparse
import os
import sys
import tempfile
import zipfile
import shutil
import pdb
import traceback
import sys

from Products.ATExtensions.ateapi import RecordField, RecordsField
from Products.CMFPlone.factory import _DEFAULT_PROFILE
from Products.CMFPlone.factory import addPloneSite
from Products.CMFPlone.utils import _createObjectByType
from AccessControl.SecurityManagement import newSecurityManager
from bika.lims.catalog import getCatalog
from Products.Archetypes import Field
from Products.CMFCore.interfaces import ITypeInformation
from Products.CMFCore.utils import getToolByName
import openpyxl


def excepthook(typ, value, tb):
import transaction

#
# def excepthook(typ, value, tb):
#     import pdb, traceback
#
    traceback.print_exception(typ, value, tb)
    pdb.pm()
    pdb.set_trace()


sys.excepthook = excepthook


# If creating a new Plone site:
default_profiles = [
    'plonetheme.classic:default',
    'plonetheme.sunburst:default',
    'plone.app.caching:default',
    'bika.lims:default',
]

export_types = [
    'Client',
    'Contact',
    'ARPriority',
    'AnalysisProfile',
    'ARTemplate',
    'AnalysisCategory',
    'AnalysisService',
    'AnalysisSpec',
    'AttachmentType',
    'BatchLabel',
    'Calculation',
    'Container',
    'ContainerType',
    'Department',
    'Instrument',
    'InstrumentCalibration',
    'InstrumentCertification',
    'InstrumentMaintenanceTask',
    'InstrumentScheduledTask',
    'InstrumentType',
    'InstrumentValidation',
    'LabContact',
    'LabProduct',
    'Manufacturer',
    'Method',
    'Preservation',
    'ReferenceDefinition',
    'SampleCondition',
    'SampleMatrix',
    'StorageLocation',
    'SamplePoint',
    'SampleType',
    'SamplingDeviation',
    'SRTemplate',
    'SubGroup',
    'Supplier',
    'SupplierContact',
    'WorksheetTemplate',
]


class Main:
    def __init__(self, args):
        self.args = args
        self.deferred_targets = []

    def __call__(self):
        """Export entire bika site
        """
        # pose as user
        self.user = app.acl_users.getUserById(self.args.username)
        newSecurityManager(None, self.user)
        # get or create portal object
        try:
            self.portal = app.unrestrictedTraverse(self.args.sitepath)
        except KeyError:
            profiles = default_profiles
            if self.args.profiles:
                profiles += list(self.args.profiles)
            addPloneSite(
                app,
                self.args.sitepath,
                title=self.args.title,
                profile_id=_DEFAULT_PROFILE,
                extension_ids=profiles,
                setup_content=True,
                default_language=self.args.language
            )
            self.portal = app.unrestrictedTraverse(self.args.sitepath)
        # Extract zipfile
        self.tempdir = tempfile.mkdtemp()
        zf = zipfile.ZipFile(self.args.inputfile, 'r')
        zf.extractall(self.tempdir)
        # Open workbook
        self.wb = openpyxl.load_workbook(
            os.path.join(self.tempdir, 'setupdata.xlsx'))
        # Import
        self.import_laboratory()
        self.import_bika_setup()
        for portal_type in export_types:
            self.import_portal_type(portal_type)
        # Remove tempdir
        shutil.rmtree(self.tempdir)

        transaction.commit()

    def get_catalog(self, portal_type):
        """grab the first catalog we are indexed in
        """
        at = getToolByName(self.portal, 'archetype_tool')
        return at.getCatalogsByType(portal_type)[0]

    def resolve_reference_ids_to_uids(self, instance, field, value):
        """Get target UIDs for any ReferenceField.
        If targets do not exist, the requirement is added to deferred_targets.
        """
        # We make an assumption here, that if there are multiple allowed
        # types, they will all be indexed in the same catalog.
        target_type = field.allowed_types \
            if isinstance(field.allowed_types, basestring) \
            else field.allowed_types[0]
        catalog = self.get_catalog(target_type)
        # The ID is what is stored in the export, so first we must grab these:
        if field.multiValued:
            # multiValued references get their values stored in a sheet
            # named after the relationship.
            ids = []
            if field.relationship[:31] not in self.wb:
                print "%s/%s: Cannot find sheet %s." % (
                    instance, field.relationship, field.relationship[:31])
                return None
            ws = self.wb[field.relationship[:31]]
            ids = []
            for rownr, row in enumerate(ws.rows):
                if rownr == 0:
                    keys = [cell.value for cell in row]
                    continue
                rowdict = dict(zip(keys, [cell.value for cell in row]))
                if rowdict['Source'] == instance.id:
                    ids.append(rowdict['Target'])
            if not ids:
                return []
            final_value = []
            for v in value:
                brain = catalog(portal_type=field.allowed_types, id=v)
                if brain:
                    final_value.append(brain.getObject())
                else:
                    self.deferred_targets.append({
                        'instance': instance,
                        'field': field,
                        'id': v
                    })
            return final_value
        else:
            brain = catalog(portal_type=field.allowed_types, id=value)
            if brain:
                return brain[0].getObject()
            else:
                self.deferred_targets.append({
                    'instance': instance,
                    'field': field,
                    'id': value
                })
        return None

    def resolve_records(self, instance, field, value):
        # RecordField and RecordsField
        # We must re-create the dict (or list of dicts) from sheet values
        ws = self.wb[value]
        matches = []
        for rownr, row in enumerate(ws.rows):
            if rownr == 0:
                keys = [cell.value for cell in row]
                continue
            rowdict = dict(zip(keys, [cell.value for cell in row]))
            if rowdict['id'] == instance.id \
                    and rowdict['field'] == field.getName():
                matches.append(rowdict)
        if type(field.default) == dict:
            return matches[0] if matches else {}
        else:
            return matches

    def set(self, instance, field, value):
        mutator = field.getMutator(instance)
        outval = self.mutate(instance, field, value)
        return mutator(outval)

    def mutate(self, instance, field, value):
        if type(value) in (int, bool):
            return value
        if isinstance(value, unicode):
            value = value.encode('utf-8')
        elif isinstance(field, RecordField):
            value = self.resolve_records(instance, field, value) \
                if value else {}
        if isinstance(field, RecordsField) or \
                (isinstance(value, basestring)
                 and value
                 and value.endswith("_values")):
            value = self.resolve_records(instance, field, value) \
                if value else []
        elif Field.IReferenceField.providedBy(field):
            value = self.resolve_reference_ids_to_uids(instance, field, value)
        elif Field.ILinesField.providedBy(field):
            value = value.splitlines() if value else ()
        # TextField implements IFileField, so we must handle it before
        # IFileField.  It does nothing to the field value.
        elif Field.ITextField.providedBy(field):
            pass
        elif value and Field.IFileField.providedBy(field):
            # XXX should not be reading entire file contents into mem.
            value = open(value).read()
        return value

    def import_laboratory(self):
        instance = self.portal.bika_setup.laboratory
        schema = instance.schema
        ws = self.wb['Laboratory']
        for row in ws.rows:
            field = schema[row[0].value]
            self.set(instance, field, row[1].value)

    def import_bika_setup(self):
        instance = self.portal.bika_setup
        schema = instance.schema
        ws = self.wb['BikaSetup']
        for row in ws.rows:
            field = schema[row[0].value]
            self.set(instance, field, row[1].value)

    def import_portal_type(self, portal_type):
        if portal_type not in self.wb:
            print "No worksheet found for type %s" % portal_type
            return None
        pt = getToolByName(self.portal, 'portal_types')
        if portal_type not in pt:
            print "Error: %s not found in portal_types." % portal_type
            return None
        fti = pt[portal_type]
        ws = self.wb[portal_type]
        keys = [cell.value for cell in ws.rows[0]]
        for rownr, row in enumerate(ws.rows[1:]):
            rowdict = dict(zip(keys, [cell.value for cell in row]))
            path = rowdict['path'].encode('utf-8').strip('/').split('/')
            del (rowdict['path'])
            uid = rowdict['uid'].encode('utf-8')
            del (rowdict['uid'])
            instance_id = rowdict['id'].encode('utf-8')
            del (rowdict['id'])
            parent = self.portal.unrestrictedTraverse(path)

            # Sampletype<->SamplePoint relations are not really good things.
            # the source setters set the back-reference directly into the
            # target object.  That's like, four extra AT objects per relation.
            # Here they get added to deferred_targets

            instance = _createObjectByType(portal_type, parent, instance_id)
            instance.unmarkCreationFlag()
            instance.reindexObject()
            for fieldname, value in rowdict.items():
                field = instance.schema[fieldname]
                self.set(instance, field, value)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Import bika setupdata created by export_bika_setup.py',
        epilog='This script is meant to be run with zopepy or bin/instance. See'
               ' http://docs.plone.org/develop/plone/misc/commandline.html for'
               ' details.'
    )
    parser.add_argument(
        '-s',
        dest='sitepath',
        required=True,
        help='full path to Plone site root.  Site will be created if it does'
             ' not already exist.')
    parser.add_argument(
        '-i',
        dest='inputfile',
        required=True,
        help='input zip file, created by the export script.')
    parser.add_argument(
        '-u',
        dest='username',
        default='admin',
        help='zope admin username (default: admin)')
    parser.add_argument(
        '-t',
        dest='title',
        help='If a new Plone site is created, this specifies the site Title.'),
    parser.add_argument(
        '-l',
        dest='language',
        default='en',
        help='If a new Plone site is created, this is the site language.'
             ' (default: en)')
    parser.add_argument(
        '-p',
        dest='profiles',
        action='append',
        help='If a new Plone site is created, this option may be used to'
             ' specify additional profiles to be activated.'),
    args, unknown = parser.parse_known_args()

    main = Main(args)
    main()
