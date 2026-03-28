import pytest
from unittest.mock import MagicMock, patch, PropertyMock
from src.services.outlook_contact_service import OutlookContactService


def make_mock_entry(name: str, address: str = None, is_dl: bool = False):
    """Helper to create a mock address entry"""
    entry = MagicMock()
    entry.Name = name
    entry.AddressEntryUserType = 1 if is_dl else 0
    entry.Address = address
    return entry


def make_mock_outlook(groups: dict):
    """
    groups: {group_name: [email1, email2, ...]}
    """
    # Build address entries
    all_entries = []
    for group_name, member_emails in groups.items():
        # Create the group entry
        dl_entry = MagicMock()
        dl_entry.Name = group_name
        dl_entry.AddressEntryUserType = 1  # Distribution list

        # Create member entries
        members = []
        for email in member_emails:
            member = MagicMock()
            member.Name = email
            member.AddressEntryUserType = 0
            member.Address = email
            members.append(member)

        # Mock members collection
        members_coll = MagicMock()
        members_coll.Count = len(members)
        members_coll.Item = lambda i, m=members: m[i - 1]
        dl_entry.Members = members_coll
        all_entries.append(dl_entry)

    # Mock address list
    addr_list = MagicMock()
    addr_list.AddressEntries.Count = len(all_entries)
    addr_list.AddressEntries.Item = lambda i, e=all_entries: e[i - 1]

    # Mock address lists collection
    addr_lists = MagicMock()
    addr_lists.Count = 1
    addr_lists.Item = lambda i: addr_list

    # Mock session and outlook
    outlook = MagicMock()
    outlook.Session.AddressLists = addr_lists
    return outlook


@pytest.fixture
def service_with_groups():
    mock_outlook = make_mock_outlook({
        '销售组': ['sales1@example.com', 'sales2@example.com'],
        'VIP客户组': ['vip1@example.com'],
    })
    return OutlookContactService(outlook_app=mock_outlook)


def test_resolve_plain_emails(service_with_groups):
    result = service_with_groups.resolve_recipients('user1@a.com;user2@b.com')
    assert result == ['user1@a.com', 'user2@b.com']


def test_resolve_group(service_with_groups):
    result = service_with_groups.resolve_recipients('销售组')
    assert 'sales1@example.com' in result
    assert 'sales2@example.com' in result


def test_resolve_mixed(service_with_groups):
    result = service_with_groups.resolve_recipients('user@x.com;销售组')
    assert 'user@x.com' in result
    assert 'sales1@example.com' in result


def test_resolve_deduplication(service_with_groups):
    result = service_with_groups.resolve_recipients('sales1@example.com;销售组')
    assert result.count('sales1@example.com') == 1


def test_resolve_empty_raises(service_with_groups):
    with pytest.raises(ValueError, match="收件人不能为空"):
        service_with_groups.resolve_recipients('')


def test_resolve_missing_group_raises(service_with_groups):
    with pytest.raises(ValueError, match="联系人组不存在"):
        service_with_groups.resolve_recipients('不存在的组')


def test_validate_group_exists(service_with_groups):
    assert service_with_groups.validate_group_exists('销售组') is True
    assert service_with_groups.validate_group_exists('不存在') is False
