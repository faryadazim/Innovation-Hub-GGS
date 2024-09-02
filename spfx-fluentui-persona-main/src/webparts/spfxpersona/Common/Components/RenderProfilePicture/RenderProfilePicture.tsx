import * as React from 'react';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import styles from './RenderProfilePicture.module.scss';

interface IProfilePicProps {
    developerName: string;
    title: string;
    getUserProfileUrl?: () => Promise<string>;
}

export function RenderProfilePicture(props: IProfilePicProps) {

    const [profileUrl, setProfileUrl] = React.useState<string>();
    let { developerName, title, getUserProfileUrl } = props;

    React.useEffect(() => {
        getUserProfileUrl().then(url => {
            setProfileUrl(url);
        });
    }, [props]);

    return (
        <div>
            <Persona
                imageUrl={profileUrl}
                text={developerName}
                secondaryText={title}
                showSecondaryText={true}
                size={PersonaSize.size32}
                imageAlt={developerName}
                styles={{ primaryText: { fontSize: '14px' }, root: { margin: '10px' } }}
            />
        </div>);
}